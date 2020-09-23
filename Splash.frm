VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form3"
   ScaleHeight     =   3255
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   480
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Loading Program Settings..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   10
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Wait (1)

Load MDIForm1
End Sub
