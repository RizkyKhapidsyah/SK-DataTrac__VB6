VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DataTrac"
   ClientHeight    =   5535
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9645
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   0   'False
      MultiLine       =   0   'False
      TextRTF         =   $"Form1.frx":000C
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      Picture         =   "Form1.frx":009A
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5160
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8114
            MinWidth        =   8114
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6218
            MinWidth        =   6218
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Forwarding"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox text4 
      Height          =   3615
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":1007
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSWinsockLib.Winsock rdest 
      Left            =   600
      Top             =   -480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock rsource 
      Left            =   120
      Top             =   -480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start Forwarding"
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Dest Host:"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Dest Port:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Source Port:"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Log"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Log"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAboutD 
         Caption         =   "Explanation of DataTrac"
      End
      Begin VB.Menu mnuHelpD 
         Caption         =   "About DataTrac"
      End
      Begin VB.Menu mnuAboutR 
         Caption         =   "About Randy Wable"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim closecount As Integer
Dim indatabuffer As String
Dim outdatabuffer As String

Private Sub Command1_Click()

If Text1 = "" Then Text1 = 23
If Text2 = "" Then
    MsgBox ("Please enter a Remote Port number.")
    Exit Sub
End If
If Text3 = "" Then
    MsgBox ("Please enter a Remote Host or IP.")
    Exit Sub
End If
'rsource.LocalPort = InputBox("Enter the Local Port:")
'rdest.RemotePort = InputBox("Enter the Remote Port:")
'rdest.RemoteHost = InputBox("Enter the Remote Host:")

rsource.LocalPort = Text1
rdest.RemotePort = Text2
rdest.RemoteHost = Text3
rsource.Listen


MsgBox ("Listening for incoming connections on port " & rsource.LocalPort)
End Sub

Private Sub Command2_Click()

rdest.Close
rsource.Close
closecount = 0
BOOLstoplisten = True

End Sub

Private Sub Command3_Click()
MsgBox (Form1.MinButton)
End Sub

Private Sub Form_Load()
closecount = 0
Text1.TabIndex = 0
StatusBar1.Panels(3) = "Listening IP: " & rsource.LocalIP
End Sub

Private Sub mnuAboutD_Click()

Load frmAbout
frmAbout.Show

End Sub

Private Sub mnuAboutR_Click()
Call ShellExecute(0&, vbNullString, "http://randy.onet.net", _
vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuExt_Click()
Unload Me
End Sub

Private Sub mnuHelpD_Click()
Load Dialog
Dialog.Show
End Sub

Private Sub mnuLoad_Click()
CommonDialog1.ShowOpen
text4.LoadFile CommonDialog1.FileName, rtfText
End Sub

Private Sub mnuSave_Click()
CommonDialog1.ShowSave
text4.SaveFile CommonDialog1.FileName, rtfText
End Sub

Private Sub Picture1_Click()
Call ShellExecute(0&, vbNullString, "http://google.com", _
vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub rdest_Close()
rsource_Close
End Sub

Private Sub rdest_DataArrival(ByVal bytesTotal As Long)
rdest.GetData indatabuffer
rsource.SendData indatabuffer
text4.Text = text4.Text & indatabuffer
indatabuffer = ""
End Sub

Private Sub rsource_Close()

    If closecount = 0 Then
    closecount = 1
    rdest.Close
    text4.Text = text4.Text & vbCrLf & "Connection Closed. " & vbCrLf & "Listening again..."
    rsource.Close
    rsource.LocalPort = Text1
    rsource.Listen
    rdest.Close
    rdest.RemotePort = Text2
    rdest.RemoteHost = Text3
    End If
    
End Sub

Private Sub rsource_ConnectionRequest(ByVal requestID As Long)
If rsource.State <> sckConnected Then rsource.Close
rsource.Accept (requestID)
rdest.Connect
closecount = 0
text4.Text = text4.Text & vbCrLf & "Connection from: " & rsource.RemoteHostIP & vbCrLf
End Sub

Private Sub rsource_DataArrival(ByVal bytesTotal As Long)
rsource.GetData outdatabuffer
rdest.SendData outdatabuffer
text4.Text = text4.Text & outdatabuffer
outdatabuffer = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then Command1_Click
End Sub

Private Sub text4_Change()
text4.SelStart = Len(text4.Text)
End Sub


Private Sub Timer1_Timer()

If rsource.State <> sckConnected Then
    StatusBar1.Panels(2) = "No Current Connections."
Else
    StatusBar1.Panels(2) = "Incoming Connection From: " & rsource.RemoteHostIP
End If

If rsource.State = sckListening Then
    StatusBar1.Panels(1) = "Listening"
ElseIf rsource.State = sckConnected Then
    StatusBar1.Panels(1) = "Connected"
Else
    StatusBar1.Panels(1) = "Not active"
End If

End Sub
