VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Not Connected"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Connection Settings"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton cmdc 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         TabIndex        =   2
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtnick 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Top             =   420
         Width           =   1815
      End
      Begin VB.TextBox txthost 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   1
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Target I.P. Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Your Nickname:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtmain 
      Height          =   5055
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8916
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0452
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   480
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   8055
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtsend 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6360
      Width           =   10335
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   240
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10800
      TabIndex        =   3
      Top             =   6360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ip Messenger Copyright Michael Howell 2002           '
'Open source all is reusuable aslong as as you email  '
' me first at Mike@Howelly.co.uk                      '
'                                                     '
'Copyright Goesto: Oliver from PSC for the menu images'
'                  Someone from PSC for Ip ping test  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''






Public Sub cmdC_Click()
Connect
End Sub
Private Sub cmdSend_Click()
On Error Resume Next

If txtsend.Text = "" Then
txtmain.Text = txtmain.Text & vbCrLf & "!!!! Server: " & txtnick.Text & " you must enter a message before you can send it !!!!!"
Exit Sub
End If

Dim strData As String

strData = txtnick.Text & ": " & txtsend.Text
txtmain.Text = txtmain.Text & vbCrLf & txtnick.Text & ": " & txtsend.Text

txtsend.Text = ""
txtsend.SetFocus

sckSend.SendData (strData)

End Sub
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
App.TaskVisible = False

sckSend.Protocol = sckUDPProtocol
cmdsend.Enabled = False

Status.Panels(1).Text = "No Connections"

strSetUser = False

MDIForm1.ComboText.Text = txtmain.SelFontName
txtsend.FontName = txtmain.SelFontName

MDIForm1.Combo2.Text = txtmain.SelFontSize
txtsend.FontSize = txtmain.SelFontSize

End Sub
Private Sub MNUBUDDY_Click()
BuddyList.Show
'Me.Hide
End Sub

Private Sub mnucolour_Click()
Cd.Flags = &H1 Or &H2
Cd.ShowColor
txtmain.SelStart = 0
txtmain.SelLength = Len(txtmain.Text)
txtmain.SelColor = Cd.Color
End Sub

Private Sub mnuConnect_Click()
Call cmdC_Click
End Sub

Private Sub mnudisconnect_Click()
sckSend.Close

cmdsend.Enabled = False
txthost.Enabled = True
txtnick.Enabled = True
txtsend.Enabled = False
mnuConnect.Enabled = True
mnuDisconnect.Enabled = False
End Sub

Private Sub mnufont_Click()

End Sub

Private Sub mnuShowIp_Click()
Form2.Show 1
End Sub

Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim TheData As String
sckSend.GetData TheData, vbString



If strSetUser = False Then
    If Left(TheData, 6) = "User: " Then
        TheData = Replace(TheData, Left(TheData, 6), "")
        Status.Panels(1) = "Connected to " & TheData
        Form1.Caption = "Connected to " & TheData
        sckSend.SendData ("User: " & txtnick.Text)
        strSetUser = True
        OppositeUser = TheData
        Timer1.Enabled = False
    Exit Sub
    End If
Exit Sub
End If

txtmain.Text = txtmain.Text & vbCrLf & TheData

txtmain.SelStart = Len(txtmain.Text)

txtsend.SelStart = Len(txtsend.Text)

Form1.WindowState = 0
End Sub
Private Sub sckSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "An error occurred in winsock!"
End
End Sub


Private Sub Timer1_Timer()
On Error Resume Next

sckSend.SendData ("User: " & txtnick.Text)
End Sub

Private Sub txthost_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Call cmdC_Click
End If
End Sub

Private Sub txtnick_Change()
If txtnick.Text <> "" Then
txthost.Enabled = True
txthost.BackColor = &H80000005
Else: txthost.BackColor = &H80000004
txthost.Enabled = False
End If
End Sub

Private Sub txtnick_KeyPress(KeyAscii As Integer)
If txtnick.Text <> "" Then
If KeyAscii = "13" Then
txthost.SetFocus
End If
End If
End Sub

Private Sub txtsend_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Call cmdSend_Click
End If
End Sub



