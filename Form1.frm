VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DSN Demo"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Test DSN Removable Yes Completely Removeable"
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "TestDSN"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete DSN"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add DSN"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database File"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Driver"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DSN Name"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Read the notes in the Function in the Module1
DSN "Add", Form1.Combo1.Text, Form1.Text1.Text, App.Path & "\Demo.mdb", Form1.Text3.Text

End Sub

Private Sub Command2_Click()
'Read the notes in the Function in the Module1
DSN "Del", Form1.Combo1.Text, Form1.Text1.Text, App.Path & "\Demo.mdb", Form1.Text3.Text

End Sub

Private Sub Form_Load()
Form1.Combo1.ListIndex = 0
Form1.Text2.Text = App.Path & "\Demo.mdb"

End Sub

