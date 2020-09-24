VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Dusk Map Editor"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7860
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   524
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6255
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   15
      Width           =   495
   End
   Begin VB.VScrollBar vscrollArray 
      Height          =   4335
      Left            =   6720
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox picArray 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   6240
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   600
      Width           =   540
   End
   Begin VB.CommandButton cmdSave 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save File To Disk"
      Top             =   2160
      Width           =   675
   End
   Begin VB.CommandButton cmdGrid 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0424
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Toggle Grid On / Off"
      Top             =   3000
      Width           =   675
   End
   Begin VB.CommandButton cmdView 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0501
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Birds Eye View"
      Top             =   4440
      Width           =   675
   End
   Begin VB.CommandButton cmdZoom 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0655
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Zoom Map Grid"
      Top             =   3720
      Width           =   675
   End
   Begin VB.CommandButton cmdNew 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":074B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Create New Map"
      Top             =   0
      Width           =   675
   End
   Begin VB.CommandButton cmdArray 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0859
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Open Image Array"
      Top             =   1440
      Width           =   675
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Open Existing Map"
      Top             =   720
      Width           =   675
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5925
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5953
            Text            =   "Image Source:  "
            TextSave        =   "Image Source:  "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2910
            MinWidth        =   2910
            Text            =   "Map Size:"
            TextSave        =   "Map Size:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2249
            MinWidth        =   2249
            Text            =   "X Position: "
            TextSave        =   "X Position: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2249
            MinWidth        =   2249
            Text            =   "Y Position:  "
            TextSave        =   "Y Position:  "
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picGrid 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   720
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
   Begin VB.HScrollBar hscrollMap 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   4335
   End
   Begin VB.VScrollBar vscrollMap 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   3975
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cdcFile 
      Left            =   1440
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PicClip.PictureClip picclipArray 
      Left            =   840
      Top             =   4320
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.Shape shapeActive 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      Height          =   525
      Left            =   6240
      Top             =   0
      Width           =   525
   End
   Begin VB.Label lblVScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "Create New Map"
      End
      Begin VB.Menu mnuMapOpen 
         Caption         =   "Open Existing Map"
      End
      Begin VB.Menu mnuArrayOpen 
         Caption         =   "Open Image Array"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Map as File"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Map"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Editor"
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "Map"
      Begin VB.Menu mnuGrid 
         Caption         =   "Toggle Grid"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "Goto Location"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert Values"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View Entire Map"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom Grid"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArray_Click()
    mnuArrayOpen_Click
End Sub

Private Sub cmdGrid_Click()
    mnuGrid_Click
End Sub

Private Sub cmdNew_Click()
    mnuNew_Click
End Sub

Private Sub cmdOpen_Click()
    mnuMapOpen_Click
End Sub

Private Sub cmdSave_Click()
    mnuSave_Click
End Sub

Private Sub cmdView_Click()
    mnuView_Click
End Sub

Private Sub cmdZoom_Click()
    mnuZoom_Click
End Sub

Private Sub Form_Load()
    vscrollArray.Enabled = False
    mnuZoom.Enabled = False
    mnuSave.Enabled = False
    mnuGrid.Enabled = False
    mnuView.Enabled = False
    mnuGoto.Enabled = False
    mnuConvert.Enabled = False
    mnuClose.Enabled = False
    cmdSave.Enabled = False
    cmdGrid.Enabled = False
    cmdZoom.Enabled = False
    cmdView.Enabled = False
    If Len(strFilePath) > 0 Then
        DoEvents
        frmOpen.Show vbModal, Me
        sbMain.Panels(2).Text = "Map Size:  " & intMapX & " x " & intMapY
        mnuSave.Enabled = True
        cmdSave.Enabled = True
        mnuGrid.Enabled = True
        cmdGrid.Enabled = True
        mnuZoom.Enabled = True
        cmdZoom.Enabled = True
        mnuConvert.Enabled = True
        mnuGoto.Enabled = True
        If picclipArray.Picture > 0 Then
            cmdView.Enabled = True
            mnuView.Enabled = True
            mnuConvert.Enabled = True
        End If
        DoEvents
        DrawMap
    End If
ErrorHandler:
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Result As String
    If boolSaved = False Then
        Result = MsgBox("Do you wish to save the current changes to " & strFilePath & "?", vbYesNoCancel, "Save Changes?")
        If Result = 2 Then Exit Sub
        If Result = 6 Then
            On Error GoTo ErrorHandler
            cdcFile.CancelError = True
            cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
            cdcFile.ShowSave
            strFilePath = cdcFile.FileName
            DoEvents
            frmSave.Show vbModal, Me
        End If
    End If
ErrorHandler:
    Unload frmMain
End Sub

Private Sub Form_Resize()
    Dim tmpWidth As Integer
    Dim tmpHeight As Integer
    Dim tmpArray As Integer
    Dim intBuffer As Integer
    intBuffer = -30
    
    tmpWidth = frmMain.ScaleWidth - 132 + intBuffer
    tmpHeight = frmMain.ScaleHeight - 50
    If tmpWidth < 50 Then picGrid.Width = 50 Else picGrid.Width = tmpWidth
    If tmpHeight < 50 Then picGrid.Height = 50 Else picGrid.Height = tmpHeight
    picGrid.Left = 48
    
    tmpArray = Int(frmMain.ScaleHeight / 33 - 2) * 33 + 3
    If tmpArray < 36 Then picArray.Height = 36 Else picArray.Height = tmpArray
    picArray.Left = frmMain.ScaleWidth - 60
    picActive.Left = frmMain.ScaleWidth - 59
    picActive.Top = 1
    shapeActive.Left = frmMain.ScaleWidth - 60
    shapeActive.Top = 0
    vscrollArray.Left = frmMain.ScaleWidth - 24
    vscrollArray.Height = picArray.Height
    vscrollArray.Top = picArray.Top - 1
    
    vscrollMap.Left = picGrid.Width + 48
    vscrollMap.Height = picGrid.Height
    hscrollMap.Left = 48
    hscrollMap.Top = picGrid.Height
    hscrollMap.Width = picGrid.Width
    
    picGrid.Refresh
    picArray.Refresh
End Sub

Private Sub hscrollMap_Change()
    intLocX = hscrollMap.Value
    lblVScroll.Visible = False
    sbMain.Panels(3).Text = "X Position:  " & hscrollMap.Value
    DoEvents
    DrawMap
End Sub

Private Sub hscrollMap_Scroll()
    lblVScroll.Left = (hscrollMap.Value + intZoom) / intMapX * (picGrid.Width - intZoom) + 48
    lblVScroll.Top = picGrid.Height + 18
    lblVScroll.Caption = hscrollMap.Value
    lblVScroll.Visible = True
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuArrayOpen_Click()
    cdcFile.CancelError = True
    On Error GoTo ErrorHandler
    cdcFile.Filter = "Gif files  (*.gif)|*.gif|Jpeg files  (*.jpg)|*.jpg|Bitmap files (*.bmp)|*.bmp|"
    cdcFile.FileName = strImagePath
    cdcFile.ShowOpen
    strImagePath = cdcFile.FileName
    picclipArray.Picture = LoadPicture(cdcFile.FileName)
    intMapClips = Int(picclipArray.Width / picclipArray.Height)
    sbMain.Panels(1).Text = "Image Source  (" & strImagePath & ")"
    vscrollArray.Enabled = True
    DoEvents
    DrawArray
    DrawMap
    If picclipArray.Picture > 0 Then
        picActive.PaintPicture picclipArray.Picture, 1, 1, 32, 32, _
            0, 0, picclipArray.Height, picclipArray.Height
    End If
    If intMapX > 0 And intMapY > 0 Then
        cmdView.Enabled = True
        mnuView.Enabled = True
        mnuConvert.Enabled = True
    End If
ErrorHandler:

End Sub

Private Sub mnuClose_Click()

    Dim Result As String
    If boolSaved = False Then
        Result = MsgBox("Do you wish to save the current changes to " & strFilePath & "?", vbYesNoCancel, "Save Changes?")
        If Result = 2 Then Exit Sub
        If Result = 6 Then
            On Error GoTo ErrorHandler
            cdcFile.CancelError = True
            cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
            cdcFile.ShowSave
            strFilePath = cdcFile.FileName
            DoEvents
            frmSave.Show vbModal, Me
        End If
    End If
    
    ReDim intMapGrid(0, 0)
    intMapX = 0
    intMapY = 0
    intLocX = 0
    intLocY = 0
    vscrollMap.Enabled = False
    hscrollMap.Enabled = False
    mnuZoom.Enabled = False
    mnuSave.Enabled = False
    mnuGrid.Enabled = False
    mnuView.Enabled = False
    mnuGoto.Enabled = False
    mnuConvert.Enabled = False
    mnuClose.Enabled = False
    cmdSave.Enabled = False
    cmdGrid.Enabled = False
    cmdZoom.Enabled = False
    cmdView.Enabled = False
    sbMain.Panels(2).Text = "Map Size:  "
    sbMain.Panels(3).Text = "X Position:  "
    sbMain.Panels(4).Text = "Y Position:  "
    boolSaved = True
    picGrid.Refresh
ErrorHandler:
    
End Sub

Private Sub mnuContents_Click()
    frmHelp.Show vbModal, Me
End Sub

Private Sub mnuConvert_Click()
    frmConvert.Show vbModal, Me
    DoEvents
    picGrid.Refresh
End Sub

Private Sub mnuExit_Click()
    Dim Result As String
    If boolSaved = False Then
        Result = MsgBox("Do you wish to save the current changes to " & strFilePath & "?", vbYesNoCancel, "Save Changes?")
        If Result = 2 Then Exit Sub
        If Result = 6 Then
            On Error GoTo ErrorHandler
            cdcFile.CancelError = True
            cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
            cdcFile.ShowSave
            strFilePath = cdcFile.FileName
            DoEvents
            frmSave.Show vbModal, Me
        End If
    End If
ErrorHandler:
    Unload Me
End Sub

Private Sub mnuGoto_Click()
    frmGoto.Show vbModal, Me
    DrawMap
End Sub

Private Sub mnuGrid_Click()
    boolMapGrid = Not boolMapGrid
    picGrid.Cls
    DrawMap
End Sub

Private Sub mnuMapOpen_Click()
    Dim Result As String
    If boolSaved = False Then
        Result = MsgBox("Do you wish to save the current changes to " & strFilePath & "?", vbYesNoCancel, "Save Changes?")
        If Result = 2 Then Exit Sub
        If Result = 6 Then
            On Error GoTo ErrorHandler
            cdcFile.CancelError = True
            cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
            cdcFile.ShowSave
            strFilePath = cdcFile.FileName
            DoEvents
            frmSave.Show vbModal, Me
        End If
    End If
    On Error GoTo ErrorHandler
    cdcFile.CancelError = True
    cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
    cdcFile.FileName = strFilePath
    cdcFile.ShowOpen
    strFilePath = cdcFile.FileName
    frmOpen.Show vbModal, Me
    If intMapX > 0 And intMapY > 0 Then
        sbMain.Panels(2).Text = "Map Size:  " & intMapX & " x " & intMapY
        mnuSave.Enabled = True
        cmdSave.Enabled = True
        mnuGrid.Enabled = True
        cmdGrid.Enabled = True
        mnuZoom.Enabled = True
        mnuClose.Enabled = True
        cmdZoom.Enabled = True
        mnuConvert.Enabled = True
        mnuGoto.Enabled = True
        If picclipArray.Picture > 0 Then
            cmdView.Enabled = True
            mnuView.Enabled = True
            mnuConvert.Enabled = True
        End If
    End If
    DoEvents
    DrawMap
ErrorHandler:

End Sub

Private Sub mnuNew_Click()
    Dim Result As String
    If boolSaved = False Then
        Result = MsgBox("Do you wish to save the current changes to " & strFilePath & "?", vbYesNoCancel, "Save Changes?")
        If Result = 2 Then Exit Sub
        If Result = 6 Then
            On Error GoTo ErrorHandler
            cdcFile.CancelError = True
            cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
            cdcFile.ShowSave
            strFilePath = cdcFile.FileName
            DoEvents
            frmSave.Show vbModal, Me
        End If
    End If
ErrorHandler:
    frmNew.Show vbModal, Me
    If intMapX > 0 And intMapY > 0 Then
        mnuSave.Enabled = True
        cmdSave.Enabled = True
        mnuGrid.Enabled = True
        cmdGrid.Enabled = True
        mnuZoom.Enabled = True
        cmdZoom.Enabled = True
        mnuGoto.Enabled = True
        mnuClose.Enabled = True
        If picclipArray.Picture > 0 Then
            cmdView.Enabled = True
            mnuView.Enabled = True
            mnuConvert.Enabled = True
        End If
    End If
End Sub

Private Sub mnuSave_Click()
    cdcFile.CancelError = True
    On Error GoTo ErrorHandler
    cdcFile.Filter = "Map files  (*.map)|*.map|All Files  (*.*)|*.*|"
    cdcFile.FileName = strFilePath
    cdcFile.ShowSave
    strFilePath = cdcFile.FileName
    DoEvents
    frmSave.Show vbModal, Me
    sbMain.Panels(2).Text = "Map Size:  " & intMapX & " x " & intMapY
    frmMain.Caption = "Dusk Map Editor:  (" & strFilePath & ")"
    mnuClose.Enabled = True
ErrorHandler:
End Sub

Private Sub mnuView_Click()
    frmColorInfo.Show vbModal, Me
End Sub

Private Sub mnuZoom_Click()
    frmZoom.Show vbModal, Me
    frmMain.hscrollMap.Max = intMapX - intZoom + 1
    frmMain.vscrollMap.Max = intMapY - intZoom + 1
    frmMain.hscrollMap.LargeChange = intZoom / 2 - 1
    frmMain.vscrollMap.LargeChange = intZoom / 2 - 1
    picGrid.Refresh
    picArray.Refresh
End Sub

Private Sub picArray_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picclipArray.Picture > 0 Then
        intCurrentArray = (Int(Y / 33) + intArrayPos)
        picActive.PaintPicture picclipArray.Picture, 1, 1, 32, 32, _
            (Int(Y / 33) + intArrayPos) * picclipArray.Height, _
            0, picclipArray.Height, picclipArray.Height
    End If
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If intMapX > 0 Then sbMain.Panels(3).Text = "X Position:  " & intLocX + Int(X / (picGrid.Width - 2) * intZoom)
    If intMapY > 0 Then sbMain.Panels(4).Text = "Y Position:  " & intLocY + Int(Y / (picGrid.Height - 2) * intZoom)
    Dim intXCurrent As Long
    Dim intYCurrent As Long
    Dim lngGridSizeY As Single
    Dim lngGridSizeX As Single
    If Button = 1 Then
        If boolMapGrid = True Then
            lngGridSizeY = 1.5
            lngGridSizeX = 1.5
        Else
            lngGridSizeY = 0
            lngGridSizeX = 0
        End If
        intXCurrent = Int(X / (picGrid.Width - 2) * intZoom) * (picGrid.Width - 2) / intZoom - 1
        intYCurrent = Int(Y / (picGrid.Height - 2) * intZoom) * (picGrid.Height - 2) / intZoom - 1
        If picclipArray.Picture > 0 And intMapY > 0 And intMapX > 0 And X > 0 And Y > 0 And Y < picGrid.Height And X < picGrid.Width Then
            picGrid.PaintPicture picclipArray.Picture, _
            intXCurrent + 1, intYCurrent + 1, _
            picGrid.Width / intZoom - lngGridSizeX, _
            picGrid.Height / intZoom - lngGridSizeY, _
            picclipArray.Height * intCurrentArray, 0, _
            picclipArray.Height - 0.6, picclipArray.Height - 0.5
            intMapGrid(Int(X / (picGrid.Width - 2) * intZoom) + intLocX, Int(Y / (picGrid.Height - 2) * intZoom) + intLocY) = intCurrentArray
        End If
    End If
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If intMapX > 0 Then sbMain.Panels(3).Text = "X Position:  " & intLocX + Int(X / (picGrid.Width - 2) * intZoom)
    If intMapY > 0 Then sbMain.Panels(4).Text = "Y Position:  " & intLocY + Int(Y / (picGrid.Height - 2) * intZoom)
    Dim intXCurrent As Long
    Dim intYCurrent As Long
    Dim lngGridSizeY As Single
    Dim lngGridSizeX As Single
    If Button = 1 Then
        If boolMapGrid = True Then
            lngGridSizeY = 1.5
            lngGridSizeX = 1.5
        Else
            lngGridSizeY = 0
            lngGridSizeX = 0
        End If
        intXCurrent = Int(X / (picGrid.Width - 2) * intZoom) * (picGrid.Width - 2) / intZoom - 1
        intYCurrent = Int(Y / (picGrid.Height - 2) * intZoom) * (picGrid.Height - 2) / intZoom - 1
        If picclipArray.Picture > 0 And intMapY > 0 And intMapX > 0 And X > 0 And Y > 0 And Y < picGrid.Height And X < picGrid.Width Then
            picGrid.PaintPicture picclipArray.Picture, _
            intXCurrent + 1, intYCurrent + 1, _
            picGrid.Width / intZoom - lngGridSizeX, _
            picGrid.Height / intZoom - lngGridSizeY, _
            picclipArray.Height * intCurrentArray, 0, _
            picclipArray.Height - 0.6, picclipArray.Height - 0.5
            intMapGrid(Int(X / (picGrid.Width - 2) * intZoom) + intLocX, Int(Y / (picGrid.Height - 2) * intZoom) + intLocY) = intCurrentArray
        End If
    End If
End Sub

Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If intMapX > 0 And intMapY > 0 Then
        mnuSave.Enabled = True
        cmdSave.Enabled = True
        boolSaved = False
    End If
    DrawMap
End Sub

Private Sub picGrid_Paint()
    DoEvents
    DrawMap
End Sub

Private Sub picArray_Paint()
    DoEvents
    DrawArray
End Sub

Private Sub vscrollArray_Change()
    intArrayPos = vscrollArray.Value
    DrawArray
End Sub

Private Sub vscrollArray_Scroll()
    intArrayPos = vscrollArray.Value
    DrawArray
End Sub

Private Sub vscrollMap_Change()
    intLocY = vscrollMap
    lblVScroll.Visible = False
    sbMain.Panels(4).Text = "Y Position:  " & vscrollMap.Value
    
    DoEvents
    DrawMap
End Sub

Private Sub vscrollMap_Scroll()
    lblVScroll.Left = picGrid.Width + 18 + 48
    lblVScroll.Top = (vscrollMap.Value + intZoom) / intMapY * (picGrid.Height - intZoom)
    lblVScroll.Caption = vscrollMap.Value
    lblVScroll.Visible = True
End Sub
