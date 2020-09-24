VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Opening and Buffering File"
   ClientHeight    =   735
   ClientLeft      =   1440
   ClientTop       =   1290
   ClientWidth     =   4095
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pbOpen 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Caption         =   "Opening and Buffering File."
      Height          =   255
      Left            =   120
      TabIndex        =   1
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
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private boolCancel As Boolean

Private Sub cmdCancel_Click()
    boolCancel = True
End Sub

Private Sub Form_Activate()

    On Error GoTo ErrorHandler

    Dim Counter As Byte
    Dim lngXSize As Long
    Dim lngYSize As Long
    Dim byteData(8) As Byte
    Dim byteInputData1 As Byte
    Dim byteInputData2 As Byte
    boolCancel = False
    
    Open strFilePath For Binary As #1
    
        For Counter = 1 To 8
            byteData(Counter) = CByte(Asc(Input(1, #1)))
            DoEvents
        Next Counter
        intMapX = byteData(3) * 256 + byteData(4) - 1
        intMapY = byteData(7) * 256 + byteData(8) - 1
        
        ReDim intMapGrid(intMapX, intMapY)
        
        pbOpen.Max = intMapX
        
        For lngXSize = 0 To intMapX
            For lngYSize = 0 To intMapY
                DoEvents
                byteInputData1 = CByte(Asc(Input(1, #1)))
                byteInputData2 = CByte(Asc(Input(1, #1)))
                intMapGrid(lngXSize, lngYSize) = (byteInputData1) * 256 + (byteInputData2)
                pbOpen.Value = lngXSize
                If boolCancel = True Then
                    lngXSize = intMapX
                    lngYSize = intMapY
                End If
            Next lngYSize
        Next lngXSize

    
    Close #1
    
    If boolCancel = True Then
        ReDim intMapGrid(0, 0)
        intMapX = 0
        intMapY = 0
    End If
    
    If intMapX < intZoom Then
        frmMain.hscrollMap.Enabled = False
        If intMapX > 0 Then intZoom = intMapX
    Else
        frmMain.hscrollMap.Enabled = True
        frmMain.hscrollMap.Max = intMapX - intZoom + 1
        frmMain.hscrollMap.LargeChange = intZoom / 2
        frmMain.mnuZoom.Enabled = True
    End If
    
    If intMapY < intZoom Then
        frmMain.vscrollMap.Enabled = False
        If intMapX > 0 Then intZoom = intMapX
    Else
        frmMain.vscrollMap.Enabled = True
        frmMain.vscrollMap.Max = intMapY - intZoom + 1
        frmMain.vscrollMap.LargeChange = intZoom / 2
        frmMain.mnuZoom.Enabled = True
    End If
    
    intLocX = 0
    intLocX = 0
    frmMain.Caption = "Dusk Map Editor:  (" & strFilePath & ")"
    boolSaved = True
    
    Unload Me
    
Exit Sub
ErrorHandler:

    MsgBox "Error while opening file:" & vbCrLf & Error & vbCrLf & vbCrLf & _
        "The file may be corrupt, or you may have selected the wrong map type."
        
    ReDim intMapLoc(0, 0)
    
    intLocX = 0
    intLocY = 0
    
    Close #1
    
    Unload Me
    
End Sub
