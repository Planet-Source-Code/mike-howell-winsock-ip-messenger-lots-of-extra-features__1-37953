VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9750
   Icon            =   "help.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label7 
      Caption         =   "Your buddys Username and IP Address"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Line Line6 
      X1              =   5280
      X2              =   7200
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Label Label6 
      Caption         =   "If you refresh your list, the program will automatically recheck the online status of everyone on your list"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Line Line5 
      X1              =   5280
      X2              =   7200
      Y1              =   1080
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Click this button to remove a buddy from your list"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Line Line4 
      X1              =   6240
      X2              =   7200
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   3600
      Y1              =   5640
      Y2              =   6000
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   4080
      Y1              =   2880
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   4080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Image Image3 
      Height          =   5685
      Left            =   3240
      Picture         =   "help.frx":0442
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3345
   End
   Begin VB.Label Label5 
      Caption         =   "Buddy List Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Click this button to add a new buddy to your list, and to automatically check their online status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "This means the user is not connected to the internet."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "This means that this user is connected to the internet, but does not necessarily  mean they are running this program"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Menu mnuclos 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub mnuclos_Click()
Unload Me
End Sub
