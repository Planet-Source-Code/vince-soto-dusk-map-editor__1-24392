VERSION 5.00
Begin VB.Form frmGoto 
   Caption         =   "Goto Location"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2070
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   2070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Y Position:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "X Position:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmGoto
End Sub

Private Sub cmdGo_Click()
    intLocX = txtX.Text
    intLocY = txtY.Text
    If intLocX < 0 Then intLocX = 0
    If intLocY < 0 Then intLocY = 0
    If intLocX > intMapX - intZoom + 1 Then intLocX = intMapX - intZoom + 1
    If intLocY > intMapY - intZoom + 1 Then intLocY = intMapY - intZoom + 1
    Unload frmGoto
End Sub
