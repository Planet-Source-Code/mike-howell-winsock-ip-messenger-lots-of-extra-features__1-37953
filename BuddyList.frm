VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form BuddyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buddy List"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3360
   Icon            =   "BuddyList.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   3360
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Buddy"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Buddy"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin ComctlLib.TreeView TV1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Your buddy list"
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7858
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuddyList.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuddyList.frx":2ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuddyList.frx":2DD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click on a buddy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   2895
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Close"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Refresh List"
   End
   Begin VB.Menu mnuchangebuddy 
      Caption         =   "Rename"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "BuddyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemNo As Integer

Private Sub Command1_Click()
Dim Nickname As String
Dim IP As String
Dim TVItem As String
Dim Statustest As String

TVItem = "T" & ItemNo
Nickname = InputBox("Please enter the user's Nickname", "Nickname")
IP = InputBox("Please Enter this user's IP." & vbCrLf & "Note: This users IP must be static", "IP")

If Nickname = "" Then
Exit Sub
End If

If IP = "" Then
Exit Sub
End If

Statustest = IP

   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize() Then
   
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
      success = Ping(Statustest, "Hello", ECHO)
      
     'display the results
      strNodeval = GetStatusCode(success)
End If
If strNodeval <> "ip success" Then

TV1.Nodes.Add "main", tvwChild, TVItem, Nickname & " - " & IP, 2, 2

Else

TV1.Nodes.Add "main", tvwChild, TVItem, Nickname & " - " & IP, 1, 1
End If

AddBuddy Nickname, IP

ItemNo = ItemNo + 1


End Sub

Private Sub Command2_Click()
On Error GoTo NodeSelect

RemoveBuddy
TV1.Nodes.Remove TV1.SelectedItem.Index

Exit Sub
NodeSelect:
MsgBox "You must first select a user to remove", vbExclamation, "Error"
End Sub

Private Sub Form_Load()
App.TaskVisible = False

ItemNo = 1

TV1.Nodes.Add , , "main", "Buddy List", 3, 3 'Create Main Parent
'    TV1.Nodes.Add "main", tvwChild, "T1", "Assman"
        
TV1.Nodes.Item(1).Expanded = True



ListBuddys
End Sub

Private Function ListBuddys()
'This function will show you how to list all of the records
'in a table

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace
Dim Max As Long
Dim RecordVal As String
'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase(App.Path & "\buddys1.mdb")
Set RS = DB.OpenRecordset("BuddyList", dbOpenTable)

'Get how manby records are in the table
Max = RS.RecordCount

If Max = "0" Then
Exit Function
End If

ItemNo = Max + 1


'Move to the begining of the file, or you can do
'RS.MoveFirst or RS.Move 1, but i prefer this
RS.Move BOF


'do the loop
For i = 1 To Max
    Dim strNodeval As String
    Dim Statustest As String
    
    'Add the data from the fields to the listbox. Notice i used
    'two different methods. One is kind of a shortcut RS!Name
    'is easy and simple. But, if you want to put in a name
    'that has a Dash '-' or use a varable like;
    'dim FieldName as String
    'FieldName = "E-Mail"
    'rs(FieldName)
    'You can not do this with the RS!FieldName Method. It dont
    'work.
RecordVal = "T" & i
Statustest = RS("IPAd")
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize() Then
   
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
      success = Ping(Statustest, "Hello", ECHO)
      
     'display the results
      strNodeval = GetStatusCode(success)
      Debug.Print strNodeval
   End If
   
   If strNodeval <> "ip success" Then
    
    TV1.Nodes.Add "main", tvwChild, RecordVal, RS("BuddyName") & " - " & RS("IPAd"), 2, 2
   Else
   
       TV1.Nodes.Add "main", tvwChild, RecordVal, RS("BuddyName") & " - " & RS("IPAd"), 1, 1
   End If
    
    
RS.MoveNext
Next i

DB.Close

End Function

Public Function AddBuddy(strNickname As String, strIpAdd As String)

Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace


Set WS = DBEngine.Workspaces(0)

Set DB = WS.OpenDatabase(App.Path & "\buddys1.mdb")

Set RS = DB.OpenRecordset("BuddyList", dbOpenTable)

RS.AddNew

RS("BuddyName") = strNickname
RS("IPAd") = strIpAdd

RS.Update

DB.Close

End Function


Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form1.WindowState = 2
End Sub

Private Sub mnuchangebuddy_Click()
Dim strInfo As String
strInfo = InputBox("Enter the name you want displayed for your selected buddy", "Change")

If strInfo = "" Then
Exit Sub
End If

Dim ItemValue As Integer
ItemValue = Replace(TV1.SelectedItem.Key, Left(TV1.SelectedItem.Key, 1), "")
ItemValue = ItemValue - 1

Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\buddys1.mdb")
'this opens a table inside the database
Set RS = DB.OpenRecordset("BuddyList", dbOpenTable)

RS.Move ItemValue

RS.Edit

RS("BuddyName") = strInfo

RS.Update
    
DB.Close

Call mnuRefresh_Click

mnuchangebuddy.Enabled = False
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuHelp_Click()
help.Show 1
End Sub

Private Sub mnuRefresh_Click()
TV1.Nodes.Clear
ItemNo = 1
TV1.Nodes.Add , , "main", "Buddy List", 3, 3 'Create Main Node
TV1.Nodes.Item(1).Expanded = True
ListBuddys
End Sub

Private Sub TV1_Click()
If TV1.SelectedItem.Key = "main" Then
mnuchangebuddy.Enabled = False
Exit Sub
End If
mnuchangebuddy.Enabled = True
End Sub

Private Sub TV1_DblClick()
If TV1.SelectedItem.Key = "main" Then
Exit Sub
End If

Dim ItemValue As Integer
ItemValue = Replace(TV1.SelectedItem.Key, Left(TV1.SelectedItem.Key, 1), "")
ItemValue = ItemValue - 1

Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\buddys1.mdb")
'this opens a table inside the database
Set RS = DB.OpenRecordset("BuddyList", dbOpenTable)

RS.Move ItemValue

Form1.txthost.Text = RS("IPAd")

DB.Close

Unload Me
Form1.Show
Form1.WindowState = 2
End Sub

Private Function RemoveBuddy()

'This function will show you how to list all of the records
'in a table

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace
Dim Max As Long
Dim RecordVal As String
'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase(App.Path & "\buddys1.mdb")
Set RS = DB.OpenRecordset("BuddyList", dbOpenTable)

Dim ItemValue As Integer
ItemValue = Replace(TV1.SelectedItem.Key, Left(TV1.SelectedItem.Key, 1), "")
ItemValue = ItemValue - 1

RS.Move ItemValue

RS.Delete

RS.Close



End Function
