VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUSR 
   Caption         =   "Neural-Net recognizing  Letters !"
   ClientHeight    =   3720
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10425
   Icon            =   "frmNN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   26
      Left            =   5145
      TabIndex        =   59
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   25
      Left            =   4725
      TabIndex        =   58
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   24
      Left            =   4305
      TabIndex        =   57
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   23
      Left            =   3885
      TabIndex        =   56
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   22
      Left            =   3465
      TabIndex        =   55
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   21
      Left            =   3045
      TabIndex        =   54
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   20
      Left            =   2625
      TabIndex        =   53
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   2205
      TabIndex        =   52
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   1785
      TabIndex        =   51
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   1365
      TabIndex        =   50
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   16
      Left            =   945
      TabIndex        =   49
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   525
      TabIndex        =   48
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   105
      TabIndex        =   47
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   5145
      TabIndex        =   46
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   4725
      TabIndex        =   45
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   4305
      TabIndex        =   44
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   3885
      TabIndex        =   43
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   3465
      TabIndex        =   42
      Top             =   10
      Width           =   330
   End
   Begin VB.TextBox TxT_AutoTRain_Wdh 
      Height          =   330
      Left            =   1470
      TabIndex        =   41
      Top             =   3255
      Width           =   645
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   3045
      TabIndex        =   40
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   2625
      TabIndex        =   39
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   2205
      TabIndex        =   38
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   1785
      TabIndex        =   37
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1365
      TabIndex        =   36
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   945
      TabIndex        =   35
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   525
      TabIndex        =   34
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_Letter 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   105
      TabIndex        =   33
      Top             =   10
      Width           =   330
   End
   Begin VB.CommandButton Cmd_AutoTrain 
      Caption         =   "Auto-Train"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   3255
      Width           =   1455
   End
   Begin VB.ListBox List_Erg 
      Columns         =   2
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   7035
      TabIndex        =   31
      Top             =   10
      Width           =   3375
   End
   Begin VB.CommandButton Make_A 
      Caption         =   "Clear Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   30
      Top             =   840
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   0
      MaxLength       =   1
      TabIndex        =   29
      Text            =   "A"
      Top             =   840
      Width           =   2010
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2415
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   27
      Top             =   1260
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   26
      Top             =   1260
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   25
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   25
      Text            =   "X"
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   24
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   24
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   23
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   23
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   22
      Left            =   2415
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   22
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   21
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "X"
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   20
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   20
      Text            =   "X"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   19
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   19
      Text            =   "X"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   18
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   18
      Text            =   "X"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   17
      Left            =   2415
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "X"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   16
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "X"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   15
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "X"
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   14
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "X"
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   13
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   2415
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "X"
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "X"
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "X"
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "X"
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   2415
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "X"
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1260
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   3045
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1260
      Width           =   285
   End
   Begin VB.TextBox Piece 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "X"
      Top             =   1260
      Width           =   285
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3780
      Top             =   1155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_Run 
      Caption         =   "->Run->"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3780
      TabIndex        =   1
      Top             =   1785
      Width           =   1185
   End
   Begin VB.CommandButton Cmd_Train 
      Caption         =   "Train"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2415
      TabIndex        =   0
      Top             =   2835
      Width           =   930
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10395
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Lbl_Erg 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5040
      TabIndex        =   28
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2415
      TabIndex        =   2
      Top             =   3255
      Width           =   915
   End
   Begin VB.Menu menu 
      Caption         =   "Main Menu"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu load 
         Caption         =   "Load Net"
      End
      Begin VB.Menu save 
         Caption         =   "Save Net"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Nicolas 'plusminus' Gramlich
' stoepsel5@gmx.de
'visit ---> http://siq.si.funpic.de and go to the 'DOWNLOAD'-Section to check more Projects !

'Credits have to remain, if you distribute this code in ANY WAY !!!!

' Neural Net recognizing Letters/or other structures, that you want the Letter to look like :D

Option Base 1
Dim PrintCode(1 To 26) As String

Private Sub about_Click()
MsgBox "Made By : Nicolas 'plus|minus' Gramlich , stoepsel5@gmx.de , visit --> http://siq.si.funpic.de ", vbOKOnly
End Sub

Private Sub Cmd_Letter_Click(Index As Integer)
MakeIT (PrintCode(Index)) 'Make the 'Picture'
Text1.Text = Chr(Index + 64)
End Sub

Private Sub Cmd_Train_Click()
'Pre-Set all OutPut-Values to ZERO
G_A = 0: G_B = 0: G_C = 0: G_D = 0: G_E = 0: G_F = 0: G_G = 0: G_H = 0: G_I = 0: G_J = 0: G_K = 0: G_L = 0: G_M = 0
G_N = 0: G_O = 0: G_P = 0: G_Q = 0: G_R = 0: G_S = 0: G_T = 0: G_U = 0: G_V = 0: G_W = 0: G_X = 0: G_Y = 0: G_Z = 0

'Cycle through the 'Pixels' and set .tag to "1" if Pixels is a "X"
For j = 1 To 25
    If Piece(j).Text = "X" Then
        Piece(j).Tag = "1"
    Else
        Piece(j).Tag = "0"
    End If
Next j

'Set The Output-Values to 1 if the Left-TExt shows its number
If Text1.Text = "A" Then G_A = 1
If Text1.Text = "B" Then G_B = 1
If Text1.Text = "C" Then G_C = 1
If Text1.Text = "D" Then G_D = 1
If Text1.Text = "E" Then G_E = 1
If Text1.Text = "F" Then G_F = 1
If Text1.Text = "G" Then G_G = 1
If Text1.Text = "H" Then G_H = 1
If Text1.Text = "I" Then G_I = 1
If Text1.Text = "J" Then G_J = 1
If Text1.Text = "K" Then G_K = 1
If Text1.Text = "L" Then G_L = 1
If Text1.Text = "M" Then G_M = 1
If Text1.Text = "N" Then G_N = 1
If Text1.Text = "O" Then G_O = 1
If Text1.Text = "P" Then G_P = 1
If Text1.Text = "Q" Then G_Q = 1
If Text1.Text = "R" Then G_R = 1
If Text1.Text = "S" Then G_S = 1
If Text1.Text = "T" Then G_T = 1
If Text1.Text = "U" Then G_U = 1
If Text1.Text = "V" Then G_V = 1
If Text1.Text = "W" Then G_W = 1
If Text1.Text = "X" Then G_X = 1
If Text1.Text = "Y" Then G_Y = 1
If Text1.Text = "Z" Then G_Z = 1


'A = InputBox("How many Itterations?", , "1500")
'If A <> vbCancel And A <> "" And IsNumeric(A) = True Then
A = 1000 'Iterations
For i = 1 To CLng(A)
    If i Mod 100 = 0 Then DoEvents
    'Insert the made Tag-Values of alle 25 'Pixels' into an Array
    myInput = Array(Piece(1).Tag, Piece(2).Tag, Piece(3).Tag, Piece(4).Tag, Piece(5).Tag, Piece(6).Tag, Piece(7).Tag, Piece(8).Tag, Piece(9).Tag, Piece(10).Tag, Piece(11).Tag, Piece(12).Tag, Piece(13).Tag, Piece(14).Tag, Piece(15).Tag, Piece(16).Tag, Piece(17).Tag, Piece(18).Tag, Piece(19).Tag, Piece(20).Tag, Piece(21).Tag, Piece(22).Tag, Piece(23).Tag, Piece(24).Tag, Piece(25).Tag)
    
    'Insert the Output thsi will be a long chain of "0" having one "1" in it, because of the 1 Letter we Set above !
    myOutput = Array(G_A, G_B, G_C, G_D, G_E, G_F, G_G, G_H, G_I, G_J, G_K, G_L, G_M, G_N, G_O, G_P, G_Q, G_R, G_S, G_T, G_U, G_V, G_W, G_X, G_Y, G_Z)  ' del
                         
    Call SupervisedTrain(myInput, myOutput) 'Run the Training
    
    Label14 = Int((i / A) * 100) & "%" 'Show the Progress
Next i

'End If
End Sub

Private Sub Cmd_Run_Click()
For j = 1 To 25
If Piece(j).Text = "X" Then
Piece(j).Tag = "1"
Else
Piece(j).Tag = "0"
End If
Next j

x = Run(Array(Piece(1).Tag, Piece(2).Tag, Piece(3).Tag, Piece(4).Tag, Piece(5).Tag, Piece(6).Tag, Piece(7).Tag, Piece(8).Tag, Piece(9).Tag, Piece(10).Tag, Piece(11).Tag, Piece(12).Tag, Piece(13).Tag, Piece(14).Tag, Piece(15).Tag, Piece(16).Tag, Piece(17).Tag, Piece(18).Tag, Piece(19).Tag, Piece(20).Tag, Piece(21).Tag, Piece(22).Tag, Piece(23).Tag, Piece(24).Tag, Piece(25).Tag)) 'RUN THE Neural Net with all 25 Pixels as Input

Lbl_Erg.Caption = "?" 'PreSet this to '?' wont change if no Output is strong enough !

List_Erg.Clear
List_Erg.AddItem "'Code' looks like ..."
For s = 1 To 26 'Cycle through the OutPut-Array
If Round(x(s), 0) = 1 Then Lbl_Erg.Caption = Chr(s + 64) 'SHOW the Letter if Output is strong enough
List_Erg.AddItem Chr(s + 64) & ": " & Round(x(s), 5) * 100 & " %" 'Add Letter to LISTBOX
If s = 13 Then List_Erg.AddItem vbNullString ' JUST HERE, to keep the style in the Listbox  =)
Next s
End Sub

Private Sub Cmd_AutoTrain_Click()
Dim Wdh As Integer
If IsNumeric(TxT_AutoTRain_Wdh.Text) = False Then Exit Sub 'Exit if no number
Wdh = TxT_AutoTRain_Wdh.Text ' Set Iterations

For i = 1 To Wdh 'Cycle through Iterations
    TxT_AutoTRain_Wdh.Text = TxT_AutoTRain_Wdh - 1 'subtract 1 from Iterations
    
    For x = 1 To 26 'cycle through all Letters
    Text1.Text = Chr(x + 64) ' Set Left TEXT to 'A' ; 'B' ; 'C' ; ....
    MakeIT (PrintCode(x)) 'Put the 01010100101-Code into a 'Picture'
    Cmd_Train_Click ' Run Train with the Settings we made
    DoEvents
    Next x
Next i
End Sub

Private Sub MakeIT(Code) ' Function to set the 01000100-Code to a 'Picture'
For x = 1 To 25
If Mid(Code, x, 1) = "1" Then
Piece(x).Text = "X"
Else
Piece(x).Text = ""
End If
Next x
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
'    'Set PRE-DEFINED Codes for every Letter
'    PrintCode(1) = "0010001110110111111110001" 'A
'    PrintCode(2) = "1110010010111001001011100" 'B
'    PrintCode(3) = "0111010000100001000001110" 'C
'    PrintCode(4) = "1110010010100101001011100" 'D
'    PrintCode(5) = "1111010000111001000011110" 'E
'    PrintCode(6) = "1111100100011100010000100" 'F
'    PrintCode(7) = "1111010000101101001011110" 'G
'    PrintCode(8) = "1000110001111111000110001" 'H
'    PrintCode(9) = "0111000100001000010001110" 'I
'    PrintCode(10) = "0111000010000100001001110" 'J
'    PrintCode(11) = "1001010100110001010010010" 'K
'    PrintCode(12) = "1000010000100001000011100" 'L
'    PrintCode(13) = "1000111011101011000110001" 'M
'    PrintCode(14) = "1000111001101011001110001" 'N
'    PrintCode(15) = "1111110001100011000111111" 'O
'    PrintCode(16) = "1110010010111001000010000" 'P
'    PrintCode(17) = "0111010001101011001001101" 'Q
'    PrintCode(18) = "1110010010111001010010010" 'R
'    PrintCode(19) = "1111110000111110000111111" 'S
'    PrintCode(20) = "1111100100001000010000100" 'T
'    PrintCode(21) = "1000110001100011000111111" 'U
'    PrintCode(22) = "1000110001100010101000100" 'V
'    PrintCode(23) = "1000110001101011010101010" 'W
'    PrintCode(24) = "1000101010001000101010001" 'X
'    PrintCode(25) = "1000101010001000100010000" 'Y
'    PrintCode(26) = "1111100010001000100011111" 'Z
Dim Maindata$
Dim FF
FF = FreeFile
Open App.Path & "\Data\Letter_Codes.txt" For Input As #FF
Do While Not EOF(1)
    Line Input #FF, Data
    Select Case Data
        Case "START Letter-Codes":
        For x = 1 To 26
            Line Input #FF, Maindata
            PrintCode(x) = Maindata
        Next x
    End Select
Loop

Close #FF

'Create the Neural Net
Call CreateNet(1.5, Array(25, 26)) '25 Input Neurons (5x5-Picture), 26 Output-Neurons for every possible Letter
End Sub

Private Sub Form_Unload(Cancel As Integer)
EraseNetwork
End Sub

Private Sub load_Click()
Dlg.Filter = "Neural nets |*.nn"
Dlg.ShowOpen
If Dlg.FileName <> "" Then
    LoadNet (Dlg.FileName)
End If
End Sub

Private Sub Make_A_Click()
For x = 1 To 25
Piece(x).Text = ""
Next x
End Sub

Private Sub Piece_Click(Index As Integer)
If Len(Piece(Index).Text) = 1 Then
Piece(Index).Text = ""
Else
Piece(Index).Text = "X"
End If

End Sub

Private Sub save_Click()
Dlg.Filter = "*.nn  Neural nets|*.nn"
Dlg.ShowSave
If Dlg.FileName <> "" Then
   SaveNet (Dlg.FileName)
End If
End Sub
