VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Map"
   ClientHeight    =   2775
   ClientLeft      =   2505
   ClientTop       =   2505
   ClientWidth     =   3120
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Map"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtYSize 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "200"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtXSize 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "200"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblTileSelect 
      Caption         =   "Select Default Tile Value:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Vertical (Y) Size:    (40 - 5000)"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Horizontal (X) Size:  (40 - 5000)"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblNew 
      Caption         =   "Specify the dimensions for the new map"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmNew
End Sub

Private Sub cmdCreate_Click()
    Dim flagsize As Boolean
    Dim MapValue As Byte
    Dim X As Single
    Dim Y As Single
    Dim Counter As Integer
    
    flagsize = False
    
    intLocX = 0
    intLocY = 0
    
    On Error GoTo ErrorHandler
    
    If Val(txtXSize) < 40 Then
        MsgBox "The horizontal (x) value you specified is too small.  Please select a value between 40 and 5000."
        txtXSize = 40
        flagsize = True
    End If
    
    If Val(txtYSize) < 40 Then
        MsgBox "The vertical (y) value you specified is too small.  Please select a value between 40 and 5000."
        txtYSize = 40
        flagsize = True
    End If
    
    If Val(txtYSize) > 5000 Then
        MsgBox "The vertical (y) value you specified is too large.  Please select a value between 40 and 5000."
        txtYSize = 5000
        flagsize = True
    End If
    
    If Val(txtXSize) > 5000 Then
        MsgBox "The horizontal (x) value you specified is too large.  Please select a value between 40 and 5000."
        txtXSize = 5000
        flagsize = True
    End If
    
    If flagsize = False Then
    
        intMapX = Val(txtXSize)
        intMapY = Val(txtYSize)
        MapValue = Val(txtValue)
        
        ReDim intMapGrid(intMapX, txtYSize)
        For X = 0 To intMapX
            For Y = 0 To intMapY
                intMapGrid(X, Y) = MapValue
            Next Y
        Next X
        With frmMain
            .mnuClose.Enabled = True
            .mnuGoto.Enabled = True
            .mnuSave.Enabled = True
            .mnuGrid.Enabled = True
            .mnuZoom.Enabled = True
            .vscrollMap.Enabled = True
            .hscrollMap.Enabled = True
            .sbMain.Panels(2).Text = "Map Size:  " & intMapX & " x " & intMapY
            .vscrollMap.Max = intMapY - intZoom + 1
            .hscrollMap.Max = intMapX - intZoom + 1
            .vscrollMap.LargeChange = intZoom / 2
            .hscrollMap.LargeChange = intZoom / 2
        End With
        
        boolSaved = False
        
        Call DrawMap
        
        Unload frmNew
        
    End If
    
Exit Sub

ErrorHandler:
    MsgBox "Unable to create map with specified parameters.  Check the values, and try again."
    Unload frmNew
End Sub

Private Sub Form_Deactivate()
    Unload frmNew
End Sub
