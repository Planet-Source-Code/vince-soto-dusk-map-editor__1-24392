VERSION 5.00
Begin VB.Form frmColorInfo 
   Caption         =   "Getting Color Information"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3300
   Icon            =   "frmColorInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Bitmap"
      End
   End
End
Attribute VB_Name = "frmColorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    DoEvents
    Dim ImageX As Integer
    Dim ImageY As Integer
    Dim Counter As Integer
    ReDim intArrayRGB(Int(frmMain.picclipArray.Width / frmMain.picclipArray.Height) - 1)
    Dim Red As Double
    Dim Green As Double
    Dim blue As Double
    
    mnuFile.Enabled = False
    picPicture.Visible = True
    picPicture.Width = frmMain.picclipArray.Height
    picPicture.Height = frmMain.picclipArray.Height
    For Counter = 0 To Int(frmMain.picclipArray.Width / frmMain.picclipArray.Height) - 1
        picPicture.PaintPicture frmMain.picclipArray.Picture _
            , 0, 0, _
            frmMain.picclipArray.Height, _
            frmMain.picclipArray.Height, _
            frmMain.picclipArray.Height * Counter, _
            0, _
            frmMain.picclipArray.Height, _
            frmMain.picclipArray.Height
        
        Red = 0
        Green = 0
        blue = 0
        For ImageX = 0 To picPicture.Width - 1
            For ImageY = 0 To picPicture.Height - 1
                DoEvents
                Red = Red + (picPicture.Point(ImageX, ImageY) Mod 65536) Mod 256 + 1
                Green = Green + (picPicture.Point(ImageX, ImageY) \ 256) Mod 256
                blue = blue + (picPicture.Point(ImageX, ImageY) \ 65536)
            Next ImageY
        Next ImageX
        avgRed = Int(Red / (ImageX) / (ImageY))
        avgGreen = Int(Green / (ImageX) / (ImageY))
        avgBlue = Int(blue / (ImageX) / (ImageY))
        intArrayRGB(Counter) = RGB(avgRed, avgGreen, avgBlue)
    Next Counter
    
    picPicture.Width = intMapX
    picPicture.Height = intMapY
    frmColorInfo.Width = frmColorInfo.Width / frmColorInfo.ScaleWidth * picPicture.Width
    frmColorInfo.Height = frmColorInfo.Height / frmColorInfo.ScaleHeight * picPicture.Height * 0.55
    frmColorInfo.Left = 0
    frmColorInfo.Top = 0

    frmColorInfo.Caption = "Drawing map, this may take several minutes!"
    DoEvents
    picPicture.AutoRedraw = True
    For ImageX = 0 To intMapX
        For ImageY = 0 To intMapY
            picPicture.PSet (ImageX, ImageY), intArrayRGB(intMapGrid(ImageX, ImageY))
        Next ImageY
    Next ImageX
    frmColorInfo.Caption = "Image complete!"
    mnuFile.Enabled = True
    DoEvents
End Sub

Private Sub mnuSave_Click()
    frmMain.cdcFile.Filter = "Bitmap File  (*.bmp)|*.bmp|"
    frmMain.cdcFile.FileName = ""
    frmMain.cdcFile.CancelError = True
    On Error GoTo ErrorHandler
    frmMain.cdcFile.ShowSave
    SavePicture picPicture.Image, frmMain.cdcFile.FileName
ErrorHandler:
End Sub

