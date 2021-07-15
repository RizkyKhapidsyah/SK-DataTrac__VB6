VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DataTrac"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Text1.Text = "DataTrac is an application that uses the Winsock component to " & _
             "bind a port from one machine to another, according to your specifications. " & vbCrLf & vbCrLf & _
             "Simply select the port you want to listen on your computer, and the destination " & _
             "computer information." & vbCrLf & vbCrLf & "Please don't abuse this software, it was not intented for malicous use."
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
