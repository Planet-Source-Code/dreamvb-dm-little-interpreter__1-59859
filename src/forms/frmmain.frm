VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   5385
      TabIndex        =   2
      Top             =   6210
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "set"
      Height          =   510
      Left            =   165
      TabIndex        =   1
      Top             =   5670
      Width           =   2505
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Text            =   "frmmain.frx":0000
      Top             =   120
      Width           =   10560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    sCommandLine = Command
End Sub

