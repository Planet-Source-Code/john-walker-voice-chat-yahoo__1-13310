VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stop_themadness"
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4875
   FillStyle       =   4  'Upward Diagonal
   Icon            =   "vcform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Win Timer"
      Height          =   510
      Left            =   2745
      TabIndex        =   11
      Top             =   4005
      Width           =   1230
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4905
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3360
            Text            =   "Welcome to STMVC chat"
            TextSave        =   "Welcome to STMVC chat"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "9:16 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/4/2000"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Room List"
      Height          =   510
      Left            =   2700
      TabIndex        =   9
      Top             =   2520
      Width           =   1725
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7965
      Top             =   180
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ICQ"
      Height          =   510
      Left            =   1440
      TabIndex        =   6
      Top             =   4005
      Width           =   1230
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "c:\rooms.txt"
      Height          =   315
      ItemData        =   "vcform.frx":0442
      Left            =   2160
      List            =   "vcform.frx":0444
      TabIndex        =   5
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   315
      Width           =   2310
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cheeta Voice Chat"
      Height          =   510
      Left            =   2700
      TabIndex        =   3
      Top             =   1260
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End VC"
      Height          =   510
      Left            =   2700
      TabIndex        =   2
      Top             =   1890
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   510
      Left            =   135
      TabIndex        =   1
      Top             =   4005
      Width           =   1230
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   4605
      Left            =   4905
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   2940
      ExtentX         =   5186
      ExtentY         =   8123
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Select the Yahoo! Chat room OR type in your own."
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   180
      TabIndex        =   8
      Top             =   765
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Enter the ID you wish to use."
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   135
      TabIndex        =   7
      Top             =   180
      Width           =   1860
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      WindowList      =   -1  'True
      Begin VB.Menu id 
         Caption         =   "IDs"
      End
      Begin VB.Menu rooms 
         Caption         =   "Rooms"
      End
      Begin VB.Menu icq 
         Caption         =   "ICQ"
      End
      Begin VB.Menu me 
         Caption         =   "Contact me"
      End
      Begin VB.Menu win 
         Caption         =   "Win Timer"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please if you use this or improve on this atleast thank me! lol
' i know this is sloppy code but hey.. i`m learning!
'if you improve this in any way please let me know
'what you have done. so i can learn more.
'i would like to make it connect to yahoo to get
'the yahoo room list so if you know how add it and let me know
' how to do it.   Stop_themadness@yahoo.com
'thanks!!!!!!!!!
'ohh and the icq part i got from here but forgot who submitted it
'sorry!!!!    but anyways  thanks to the programer of the icq pager!!!
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        Command3_Click
    End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()


WebBrowser1.Navigate "about:blank"
WebBrowser1.Visible = False
Timer1.Enabled = False
Form1.Width = 4680
End Sub

Private Sub Command3_Click()
On Error Resume Next


        
   
    
Form1.Width = 8100
WebBrowser1.Visible = True
On Error Resume Next
If Text1.Text = "" Then
Text1.Text = "%20"
End If

If Combo1.Text = "" Then
GoTo start:
End If
Open "c:\rooms.stm" For Append As #1    ' Open file for output.
Write #1, Combo1.Text    ' Write comma delimited data .
Write #1,   ' Write blank line.
Close #1    ' Close file.
start:
     WebBrowser1.Navigate "http://login.cheetachat.net/cvoice/voice.php3?k=chya%5F&c=" + (Form1.Combo1.Text) + "&u=" + (Form1.Text1.Text)
 Text1.Text = ""
Timer1.Enabled = True
 

End Sub
Private Sub Command4_Click()
FormMain.Visible = True
End Sub

Private Sub Command5_Click()
On Error Resume Next
Kill ("c:\rooms.stm")
Combo1.Clear
End Sub

Private Sub Command6_Click()
Dim lngReturn As Long
lngReturn = GetTickCount()
MsgBox ("Windows has been running for " & (lngReturn / 1000) & " seconds."), vbOKOnly, "Win Timer"
End Sub

Private Sub Form_Load()
App.TaskVisible = False
If App.PrevInstance = True Then
 MsgBox "This program is already running !", vbCritical
End
End If

Dim record As String
Open "c:\rooms.stm" For Append As #1
Write #1, "The Tiki Lounge:1"
Close #1
Open "c:\rooms.stm" For Input As #1
While Not EOF(1)
Input #1, record
Combo1.AddItem record
Wend
Close #1


Form1.Width = 4680
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub icq_Click()
MsgBox "OOOOOK with this you can send pager messages ** warning** the recipient WILL get your IP address in the message. The IP address is inserted by ICQ on their server and I can't do anything about it. I just thought this would be nice to add in the program.", vbExclamation, "ICQ"
End Sub

Private Sub id_Click()
MsgBox "You are able to login to Cheeta VC with any ID you wish to use. You may use a blank ID or an ID that is already in VC. Use that ID you always wanted but some ass hole already have or just make up a cool one!", vbOKOnly, "ID"
End Sub

Private Sub me_Click()
MsgBox "Thank you for trying my program. You can contact me for bug reports or input on this or any other program I have made at Stop_themadness@yahoo.com All input is welcome.", vbOKOnly, "Thank You"
End Sub

Private Sub rooms_Click()
MsgBox "You may use any yahoo room name + # or just make up your own! <As Long as your friends know the name of the room they can join you>. Don't for get to type in the name correctly, for instance this is how you would type in    The Tiki Lounge:1    Notice the pull down box? this is like a room history it remembers the rooms you have inputed and you can clear the list with the Clear room list button", vbOKOnly, "Rooms"
End Sub

Private Sub win_Click()
MsgBox "OHH COME ON! YOU KNOW WHAT THAT IS!!!!!", vbOKOnly, "Win Timer"
End Sub
