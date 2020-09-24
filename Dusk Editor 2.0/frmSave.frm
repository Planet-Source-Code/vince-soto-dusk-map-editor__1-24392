VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Opening and Buffering File"
   ClientHeight    =   735
   ClientLeft      =   1440
   ClientTop       =   1290
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbSave 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000015&
      X1              =   272
      X2              =   272
      Y1              =   1
      Y2              =   48
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   273
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label lblProgress 
      Caption         =   "Writing Map File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   271
      X2              =   271
      Y1              =   1
      Y2              =   48
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   1
      X2              =   271
      Y1              =   47
      Y2              =   47
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   1
      X2              =   1
      Y1              =   46
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   1
      X2              =   272
      Y1              =   1
      Y2              =   1
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolCancel As Boolean

Private Sub cmdCancel_Click()
    boolCancel = True
End Sub

Private Sub Form_Activate()
    Dim intX As Integer
    Dim intY As Integer
    
    Open strFilePath For Output As #1
    Print #1, Chr(0);
    Print #1, Chr(0);
    Print #1, Chr(Int(intMapX / 256));
    Print #1, Chr(intMapX Mod 256 + 1);
    Print #1, Chr(0);
    Print #1, Chr(0);
    Print #1, Chr(Int(intMapY / 256));
    Print #1, Chr(intMapY Mod 256 + 1);
    
    On Error GoTo ErrorHandler
    frmSave.Caption = "Writing and Saving Map"
    frmSave.Refresh
    
    pbSave.Max = intMapX
    
    For intX = 0 To intMapX
        For intY = 0 To intMapY
            DoEvents
            If boolCancel Then
                Exit For
                GoTo ErrorHandler
            End If
            Print #1, Chr(Int(intMapGrid(intX, intY) / 256));
            Print #1, Chr(intMapGrid(intX, intY) Mod 256);
            pbSave.Value = intX
        Next intY
    Next intX
    Close #1
    boolSaved = True
    
    Unload frmSave
Exit Sub
ErrorHandler:
    MsgBox "There was an error saving the map.  Changes may have not been saved.  Please try again."
    Close #1
    
    boolSaved = False
    
    Unload frmSave
End Sub
