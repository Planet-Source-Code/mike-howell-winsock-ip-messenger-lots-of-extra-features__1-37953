VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   " Messenger"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7065
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   5640
      Top             =   960
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   2400
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   741
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   7065
      _CBHeight       =   420
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5520
         Top             =   0
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Top             =   60
         Width           =   975
      End
      Begin VB.ComboBox ComboText 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   60
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Height          =   320
         Left            =   480
         Picture         =   "MDIForm1.frx":0452
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   320
         Left            =   120
         Picture         =   "MDIForm1.frx":059C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   375
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":06E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuShowIp 
         Caption         =   "What's my IP?"
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuBuddy 
      Caption         =   "Buddy List"
   End
   Begin VB.Menu systraymenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnushow 
         Caption         =   "Restore Program"
      End
      Begin VB.Menu mnushowbuddy 
         Caption         =   "Buddy List"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitsystray 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo2_Click()
Form1.txtmain.SelStart = 0
Form1.txtmain.SelLength = Len(Form1.txtmain.Text)
Form1.txtmain.SelFontSize = Combo2.Text
Form1.txtsend.FontSize = Combo2.Text
Form1.txtmain.SelLength = "0"

Form1.txtsend.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Form1.txtmain.SelStart = 0
Form1.txtmain.SelLength = Len(Form1.txtmain.Text)
Form1.txtmain.SelFontSize = ComboText.Text
Form1.txtsend.FontSize = Combo2.Text
Form1.txtmain.SelLength = "0"
Form1.txtsend.SetFocus
End If
End Sub



Private Sub ComboText_Click()
If ComboText.Text = "" Then
Exit Sub
End If

Form1.txtmain.SelStart = 0
Form1.txtmain.SelLength = Len(Form1.txtmain.Text)
Form1.txtmain.SelFontName = ComboText.Text
Form1.txtsend.FontName = ComboText.Text
Form1.txtmain.SelLength = "0"
Form1.txtsend.SetFocus
End Sub

Private Sub Command1_Click()
Cd.ShowSave
Cd.Filter = "Text Files (*.txt)|*.txt|"


If Cd.FileName = "" Then
    Exit Sub
End If

Open Cd.FileName For Output As #1 'create/open the file
Write #1, Form1.txtmain.Text 'write the Encrypted txt to the file
Close #1 'close the file

End Sub

Private Sub Command4_Click()
Load Form1
End Sub

Private Sub MDIForm_Load()
Name App.Path & "\buddys1.ipdb" As App.Path & "\buddys1.mdb"

    For i = 1 To Screen.FontCount
        ComboText.AddItem Screen.Fonts(i - 1)
    Next i

Unload Splash

Load Form1

Combo2.AddItem "6"
Combo2.AddItem "8"
Combo2.AddItem "10"
Combo2.AddItem "12"
Combo2.AddItem "14"
Combo2.AddItem "16"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MDIForm1.WindowState = 1 Then


Select Case X
        Case 7755:   'Right Click
            PopupMenu systraymenu  'The systray menu works the same as
                            'clicking file on the form. Anything
                            'you can do with a menu on the form
                            'you can do in the systray.
            
        
        Case 7725:    'Dbl Left Click
            MDIForm1.Show
            MDIForm1.WindowState = 2
    End Select
End If
End Sub

Private Sub MDIForm_Terminate()
'Name App.Path & "\buddys1.mdb" As App.Path & "\buddys1.ipdb"

End Sub

Private Sub MNUBUDDY_Click()
BuddyList.Show
BuddyList.SetFocus
Form1.Hide
End Sub

Private Sub mnucolour_Click()
Cd.Flags = &H1 Or &H2
Cd.ShowColor
txtmain.SelStart = 0
txtmain.SelLength = Len(txtmain.Text)
txtmain.SelColor = Cd.Color
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, try
Name App.Path & "\buddys1.mdb" As App.Path & "\buddys1.ipdb"
End Sub

Private Sub mnuConnect_Click()
Connect
End Sub

Private Sub mnudisconnect_Click()
Form1.sckSend.Close
Call GUIDisconnected
End Sub

Private Sub mnufont_Click()
Cd.Flags = &H1
Cd.ShowFont
txtmain.SelStart = 0
txtmain.SelLength = Len(txtmain.Text)
txtmain.SelFontName = Cd.FontName
End Sub

Private Sub mnuExitsystray_Click()
End
End Sub

Private Sub mnushow_Click()
            MDIForm1.Show
            MDIForm1.WindowState = 2
End Sub

Private Sub mnushowbuddy_Click()
            MDIForm1.Show
            MDIForm1.WindowState = 2
BuddyList.Show
BuddyList.SetFocus
Form1.Hide
End Sub

Private Sub mnuShowIp_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub Timer1_Timer()

If Me.WindowState = 1 Then
    Me.Hide
    'This gets Loaded when your form starts
try.cbSize = Len(try)
try.hwnd = Me.hwnd
try.uId = vbNull
try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
try.uCallBackMessage = WM_MOUSEMOVE

'To Change the Icon Displayed in the systray
'Change the Forms Icon
'This uses whatever Icon the Form Displays
try.hIcon = Me.Icon

'Tool Tip
try.szTip = "IP Messenger" & vbNullChar

Call Shell_NotifyIcon(NIM_ADD, try)
Call Shell_NotifyIcon(NIM_MODIFY, try)

'If u just want the systay icon to appear at start Hide the Form
'Me.Hide

End If

If Me.WindowState = 2 Then
Shell_NotifyIcon NIM_DELETE, try
End If

End Sub


Private Sub Timer2_Timer()
strnewchat = Form1.txtmain.Text
If MDIForm1.WindowState = 1 Then
    If strnewchat <> stroldchat Then
        MDIForm1.WindowState = 2
        MDIForm1.Show
    End If
End If
stroldchat = Form1.txtmain.Text
strnewchat = ""
End Sub
