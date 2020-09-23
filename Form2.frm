VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   780
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   2490
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
Form1.WindowState = 2
End Sub

Private Sub Form_Load()
App.TaskVisible = False

Text1.Text = Winsock1.LocalIP
Me.Top = Form1.Top + 600
Me.Left = Form1.Left + 600
End Sub

