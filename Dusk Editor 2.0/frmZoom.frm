VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom Settings"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.Slider sliderZoom 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      Min             =   10
      Max             =   40
      SelStart        =   20
      TickFrequency   =   10
      Value           =   20
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      Caption         =   "40x"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      Caption         =   "30x"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      Caption         =   "10x"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      Caption         =   "20x"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblZoom 
      Caption         =   "Zoom:  20 x 20"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
    intZoom = sliderZoom.Value
    If intLocX > intMapX - intZoom - 1 Then intLocX = intMapX - intZoom - 1
    If intLocY > intMapY - intZoom - 1 Then intLocY = intMapY - intZoom - 1
    Unload frmZoom
End Sub

Private Sub cmdCancel_Click()
    Unload frmZoom
End Sub

Private Sub Form_Load()
    sliderZoom.Value = intZoom
End Sub

Private Sub sliderZoom_Change()
    lblZoom.Caption = "Zoom:  " & sliderZoom.Value & _
        " x " & sliderZoom.Value
End Sub
