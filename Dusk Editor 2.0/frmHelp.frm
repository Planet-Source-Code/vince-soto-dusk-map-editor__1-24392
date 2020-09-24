VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help Documentation"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   4335
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload frmHelp
End Sub

Private Sub Form_Activate()
    Dim strHelp As String
    Dim strAdd As String
    On Error GoTo ErrorHandler
    Open (App.Path & "\Readme.txt") For Input As #1
    Do Until EOF(1)
        Line Input #1, strAdd
        strHelp = strHelp + strAdd + vbCrLf
    Loop
    Close #1
    txtHelp.Text = strHelp
Exit Sub
ErrorHandler:
    MsgBox "The help file, readme.txt could not be located." & vbCrLf & "You may need to reinstall the program for this functionality.", vbCritical
    Unload frmHelp
    Close #1
End Sub
