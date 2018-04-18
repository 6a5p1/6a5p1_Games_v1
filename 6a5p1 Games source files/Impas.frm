VERSION 5.00
Begin VB.Form Impasse 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "6a5p1 IMPASSE"
   ClientHeight    =   6525
   ClientLeft      =   4185
   ClientTop       =   1905
   ClientWidth     =   7200
   Icon            =   "Impas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Impas.frx":030A
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Impas.frx":0614
   ScaleHeight     =   6525
   ScaleMode       =   0  'User
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6720
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6720
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6720
      Top             =   2760
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
      Left            =   6135
      MouseIcon       =   "Impas.frx":18B5
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   75
      Width           =   255
   End
   Begin VB.Label Changepieces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change pieces"
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
      MouseIcon       =   "Impas.frx":1BBF
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   495
      Index           =   8
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label12 
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
      MouseIcon       =   "Impas.frx":1EC9
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape12 
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
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   -120
      X2              =   7560
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   4
      X1              =   6000
      X2              =   6000
      Y1              =   3120
      Y2              =   5160
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   65
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   64
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   63
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   62
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   61
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   60
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   6120
      TabIndex        =   59
      Top             =   4800
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   6000
      X2              =   6120
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   5880
      X2              =   6000
      Y1              =   3360
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   58
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   5520
      X2              =   5760
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   5520
      X2              =   5760
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   3720
      X2              =   5760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Steps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S t e p s"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2640
      MouseIcon       =   "Impas.frx":21D3
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label AboutGame 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5280
      MouseIcon       =   "Impas.frx":24DD
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   720
      Width           =   975
   End
   Begin VB.Label HelpGame 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5280
      MouseIcon       =   "Impas.frx":27E7
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   480
      Width           =   975
   End
   Begin VB.Label labelwarning 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "I N C O R R E C T "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label pasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5040
      MouseIcon       =   "Impas.frx":2AF1
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label pasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4200
      MouseIcon       =   "Impas.frx":2DFB
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label pasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3240
      MouseIcon       =   "Impas.frx":3105
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label pasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      MouseIcon       =   "Impas.frx":340F
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label pasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      MouseIcon       =   "Impas.frx":3719
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   35
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   34
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   33
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   32
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   31
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   30
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   29
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   28
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   27
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   26
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   25
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   24
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   23
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   22
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   21
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   20
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   19
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   18
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1                                          Impasse"
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
      TabIndex        =   30
      Top             =   5880
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   17
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   16
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   15
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   14
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   13
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   12
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   11
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   10
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   9
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   2
      X1              =   960
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label newgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      MouseIcon       =   "Impas.frx":3A23
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   480
      Width           =   735
   End
   Begin VB.Label exitgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   6360
      MouseIcon       =   "Impas.frx":3D2D
      MousePointer    =   99  'Custom
      TabIndex        =   19
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
      MouseIcon       =   "Impas.frx":4037
      MousePointer    =   99  'Custom
      TabIndex        =   18
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
      Caption         =   "6a5p1 IMPASSE v1.0"
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
      Left            =   2280
      MouseIcon       =   "Impas.frx":4341
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   60
      Width           =   2655
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
      TabIndex        =   13
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
      MouseIcon       =   "Impas.frx":464B
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   2
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   1
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   0
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   5
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   4
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   3
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   8
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   7
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   495
      Index           =   6
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label HelpGame1 
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
      Left            =   5400
      TabIndex        =   15
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
      MouseIcon       =   "Impas.frx":4955
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label aboutgame1 
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
      Left            =   5400
      TabIndex        =   12
      Top             =   720
      Width           =   735
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
      Left            =   2040
      MouseIcon       =   "Impas.frx":4C5F
      MousePointer    =   99  'Custom
      TabIndex        =   11
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
      Left            =   960
      MouseIcon       =   "Impas.frx":4F69
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   3000
      MouseIcon       =   "Impas.frx":5273
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2280
      MouseIcon       =   "Impas.frx":557D
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1560
      MouseIcon       =   "Impas.frx":5887
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5160
      MouseIcon       =   "Impas.frx":5B91
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4440
      MouseIcon       =   "Impas.frx":5E9B
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3720
      MouseIcon       =   "Impas.frx":61A5
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3000
      MouseIcon       =   "Impas.frx":64AF
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2280
      MouseIcon       =   "Impas.frx":67B9
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1680
      Width           =   615
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
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
      MouseIcon       =   "Impas.frx":6AC3
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1680
      Width           =   615
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
      Index           =   1
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
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
      Left            =   2040
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
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   5160
      MouseIcon       =   "Impas.frx":6DCD
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   4440
      MouseIcon       =   "Impas.frx":70D7
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3720
      MouseIcon       =   "Impas.frx":73E1
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   5160
      MouseIcon       =   "Impas.frx":76EB
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   4440
      MouseIcon       =   "Impas.frx":79F5
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3720
      MouseIcon       =   "Impas.frx":7CFF
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3000
      MouseIcon       =   "Impas.frx":8009
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   2280
      MouseIcon       =   "Impas.frx":8313
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   1560
      MouseIcon       =   "Impas.frx":861D
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   5160
      MouseIcon       =   "Impas.frx":8927
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   4440
      MouseIcon       =   "Impas.frx":8C31
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   3720
      MouseIcon       =   "Impas.frx":8F3B
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   3000
      MouseIcon       =   "Impas.frx":9245
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   2280
      MouseIcon       =   "Impas.frx":954F
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   1560
      MouseIcon       =   "Impas.frx":9859
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   5160
      MouseIcon       =   "Impas.frx":9B63
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   4440
      MouseIcon       =   "Impas.frx":9E6D
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   3720
      MouseIcon       =   "Impas.frx":A177
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   3000
      MouseIcon       =   "Impas.frx":A481
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   2280
      MouseIcon       =   "Impas.frx":A78B
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   1560
      MouseIcon       =   "Impas.frx":AA95
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   5160
      MouseIcon       =   "Impas.frx":AD9F
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   4440
      MouseIcon       =   "Impas.frx":B0A9
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3720
      MouseIcon       =   "Impas.frx":B3B3
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   3000
      MouseIcon       =   "Impas.frx":B6BD
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2280
      MouseIcon       =   "Impas.frx":B9C7
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   1560
      MouseIcon       =   "Impas.frx":BCD1
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   3480
      Width           =   615
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
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   0
      Picture         =   "Impas.frx":BFDB
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7215
   End
   Begin VB.Image Image5 
      Height          =   6135
      Left            =   0
      Picture         =   "Impas.frx":3E9E2
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image Image4 
      Height          =   6120
      Left            =   0
      Picture         =   "Impas.frx":4C27C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.Image Image3 
      Height          =   6135
      Left            =   0
      Picture         =   "Impas.frx":8082E
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   6135
      Left            =   0
      Picture         =   "Impas.frx":8851E
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
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
Attribute VB_Name = "Impasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teren(36) As Integer
Dim SLCTD As Boolean
Dim cm As Integer
Dim a(6, 6) As Integer
Dim i, j, k, p As Integer
Dim io, jo, po, depo, dep, s, tt, v As Integer
Dim piece1, piece2 As String

Private Sub aboutgame_Click()
WhichAbout = 3: GameAbout.Show
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
Private Sub Redraw()
For X = 0 To 35
    If t(X).Caption <> "" Then t(X).Caption = piece1
Next X
End Sub

Private Sub Changepieces_Click()
If piece1 = "|" Then
    piece1 = "l"
    piece2 = "S"
ElseIf piece1 = "l" Then
    piece1 = "["
    piece2 = "m"
ElseIf piece1 = "[" Then
    piece1 = "{"
    piece2 = ">"
ElseIf piece1 = "{" Then
    piece1 = "J"
    piece2 = "L"
ElseIf piece1 = "J" Then
    piece1 = "|"
    piece2 = "X"
End If
Redraw
End Sub

Private Sub changesound_Click()
SoundChooser.Show
End Sub
Private Sub changesquares_Click()
If t(0).BackColor = &HC0FFC0 Then
  For m = 0 To 35
  t(m).BackColor = &HFFFFC0
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &HFFFF80
  Next m
  labelwarning.BackColor = &HFFFFC0
  For m = 1 To 5
  pasi(m).BackColor = &HFFFF80
  Next m
ElseIf t(0).BackColor = &HFFFFC0 Then
  For m = 0 To 35
  t(m).BackColor = &HFFC0FF
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &HFF80FF
  Next m
  labelwarning.BackColor = &HFFC0FF
  For m = 1 To 5
  pasi(m).BackColor = &HFF80FF
  Next m
ElseIf t(0).BackColor = &HFFC0FF Then
  For m = 0 To 35
  t(m).BackColor = &HC0E0FF
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &H80C0FF
  Next m
  labelwarning.BackColor = &HC0E0FF
  For m = 1 To 5
  pasi(m).BackColor = &H80C0FF
  Next m
ElseIf t(0).BackColor = &HC0E0FF Then
  For m = 0 To 35
  t(m).BackColor = vbWhite
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &HE0E0E0
  Next m
  labelwarning.BackColor = vbWhite
  For m = 1 To 5
  pasi(m).BackColor = &HE0E0E0
  Next m
ElseIf t(0).BackColor = vbWhite Then
  For m = 0 To 35
  t(m).BackColor = &HC0FFFF
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &H80FFFF
  Next m
  labelwarning.BackColor = &HC0FFFF
  For m = 1 To 5
  pasi(m).BackColor = &H80FFFF
  Next m
ElseIf t(0).BackColor = &HC0FFFF Then
  For m = 0 To 35
  t(m).BackColor = &HC0C0FF
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &H8080FF
  Next m
  labelwarning.BackColor = &HC0C0FF
  For m = 1 To 5
  pasi(m).BackColor = &H8080FF
  Next m
ElseIf t(0).BackColor = &HC0C0FF Then
  For m = 0 To 35
  t(m).BackColor = &HFFC0C0
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &HFF8080
  Next m
  labelwarning.BackColor = &HFFC0C0
  For m = 1 To 5
  pasi(m).BackColor = &HFF8080
  Next m
ElseIf t(0).BackColor = &HFFC0C0 Then
  For m = 0 To 35
  t(m).BackColor = &HC0FFC0
  Next m
  For m = 0 To 8
  Shape3(m).BackColor = &H80FF80
  Next m
  labelwarning.BackColor = &HC0FFC0
  For m = 1 To 5
  pasi(m).BackColor = &H80FF80
  Next m
End If
End Sub

Private Sub exitgame_Click()
WhichQuit = 2: GameQuit.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then WhichQuit = 2: GameQuit.Show
End Sub

Private Sub Form_Load()
piece1 = "|"
piece2 = "X"
If Language = 1 Then
    newgame2.Caption = "Joc nou"
    changebackground.Caption = "Schimbã fundal"
    changesquares.Caption = "Schimbã temã"
    changesound.Caption = "Sunet"
    Mute.Caption = "Oprit"
    aboutgame1.Caption = "Despre"
    HelpGame1.Caption = "Ajutor"
    exitgame2.Caption = "Ieºire"
    Changepieces.Caption = "Schimbã piese"
    Steps.Caption = "P a º i"
    Label11.Caption = "J"
    Label10.Caption = "U"
    Label9.Caption = "C"
    Label8.Caption = "Ã"
    Label7.Caption = "T"
    Label6.Caption = "O"
    Label5.Caption = "R"
    labelwarning.Caption = "I N C O R E C T "
    labelwarning.Height = 3615
End If
Unload Game
newgame_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub

Private Sub HelpGame_Click()
WhichHelp = 2: GameHelp.Show
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub

Private Sub Label12_Click()
WindowState = 1
Unload GameWell
Unload GameAbout
Unload GameHelp
Unload GameQuit
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
WhichAbout = 3: GameAbout.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.FontSize = 9 Then Label3.FontSize = 11
End Sub

Private Sub Label4_Click()
WhichQuit = 2: GameQuit.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HC0& Then Shape4.BackColor = &HFF&
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape12.BackColor = &HC0& Then Shape12.BackColor = &HFF&
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape13.BackColor = &HC0& Then Shape13.BackColor = &HFF&
End Sub
Private Sub Mute_Click()
SoundChooser.muzica.Command = "close"
End Sub

Private Sub newgame_Click()
For i = 1 To 6
For j = 1 To 6
a(i, j) = 0
Next j
Next i
For i = 0 To 35
t(i).Caption = ""
Next i
For i = 4 To 6
For j = 1 To 3
a(i, j) = 1
t(6 * (i - 1) + j - 1).Caption = piece1
Next j
Next i
Randomize
cm = Int(Rnd * 2)
If cm = 1 Then randcomp
End Sub

Private Sub randmeu()
For k = i - 1 To i - p Step -1
    If a(k, j) = 1 Then labelwarning.Visible = True: Exit Sub
Next k
If a(i, j) = 0 Then labelwarning.Visible = True: Exit Sub
a(i, j) = 0: a(i - p, j) = 1
t(6 * (i - 1) + j - 1).Caption = piece2
t(6 * (i - p - 1) + j - 1).Caption = piece2
Timer2.Enabled = True
End Sub
Private Sub randcomp()
SLCTD = False
io = 1: jo = 1: po = 0
depo = 1000
For i = 1 To 6
    For j = 1 To 5
        If a(i, j) = 0 Or a(i, j + 1) = 1 Then GoTo 480
        For k = j + 1 To 6
            If a(i, k) = 1 Then GoTo 480
            a(i, j) = 0: a(i, k) = 1
            dep = 0
            For tt = 2 To 6
                For s = 1 To 6
                    If a(tt, s) = 0 Then GoTo 432
                    For v = tt - 1 To 1 Step -1
                        If a(v, s) = 1 Then GoTo 432
                        dep = dep + 1
                    Next v
432             Next s
            Next tt
        a(i, j) = 1: a(i, k) = 0
        If depo <= dep Then GoTo 480
        depo = dep: io = i: jo = j: po = k - j
        If depo = 0 Then GoTo 500
        Next k
480 Next j
Next i
500 a(io, jo) = 0: a(io, jo + po) = 1
t(6 * (io - 1) + jo - 1).Caption = piece2
t(6 * (io - 1) + jo + po - 1).Caption = piece2
Timer3.Enabled = True
End Sub

Private Sub pasi_Click(indecs As Integer)
If indecs < 1 Or i - indecs < 1 Or SLCTD = False Then
    labelwarning.Visible = True: Exit Sub
Else
    labelwarning.Visible = False: p = indecs: randmeu: Exit Sub
End If
End Sub

Private Sub t_Click(Index As Integer)
SLCTD = False
For X = 0 To 35
If t(X).Caption = piece2 Then t(X).Caption = piece1
Next X
If t(Index).Caption <> "" Then
    t(Index).Caption = piece2
    SLCTD = True
    j = Index - 6 * (Int(Index / 6)) + 1
    i = Int(Index / 6) + 1
End If
End Sub

Private Sub Timer1_Timer()
randcomp
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
t(6 * (i - 1) + j - 1).Caption = ""
t(6 * (i - p - 1) + j - 1).Caption = piece1
For i = 1 To 6
    For j = 1 To 5
        If a(i, j) = 1 And a(i, j + 1) = 0 Then Timer1.Enabled = True: Timer2.Enabled = False: Exit Sub
    Next j
Next i
WhichWell = 0
GameWell.Show
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
t(6 * (io - 1) + jo - 1).Caption = ""
t(6 * (io - 1) + jo + po - 1).Caption = piece1
For i = 2 To 6
    For j = 1 To 6
        If a(i, j) = 1 And a(i - 1, j) = 0 Then Timer3.Enabled = False: Exit Sub
    Next j
Next i
WhichWell = 2
GameWell.Show
Timer3.Enabled = False
End Sub
