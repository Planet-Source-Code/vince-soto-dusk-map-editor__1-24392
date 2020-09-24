VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Dusk Map Editor"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   240
      Top             =   1680
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Send comments to VJSoto@Yahoo.com"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "06-23-2001"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "v 2.0.0"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "By: Vincent Soto"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Dusk Map Editor"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bytColor As Byte
Dim colorUp As Integer

Private Sub cmdOk_Click()
    Unload frmAbout
End Sub

Private Sub Form_Activate()
    bytColor = 0
    colorUp = 5
End Sub

Private Sub Timer1_Timer()
        DoEvents
        bytColor = bytColor + colorUp
        If bytColor > 250 Then
            bytColor = 250
            colorUp = -1
        End If
        If bytColor < 5 Then
            bytColor = 5
            colorUp = 1
        End If
        Label1(0).BackColor = RGB(bytColor, 0, 255 - bytColor)
        Label1(0).ForeColor = RGB(bytColor, 255, 255 - bytColor)
        Label1(0).Left = -bytColor * 7
End Sub
