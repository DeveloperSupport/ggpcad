VERSION 5.00
Begin VB.Form SC 
   Caption         =   "GoGreen PC Art Designer Scratch Pad"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   LinkTopic       =   "SC"
   ScaleHeight     =   5775
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Left            =   5760
      Top             =   4200
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "GoGreenPCartDesignScratchPad.frx":0000
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'scratch form
End Sub
