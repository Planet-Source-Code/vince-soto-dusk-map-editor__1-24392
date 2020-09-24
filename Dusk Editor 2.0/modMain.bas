Attribute VB_Name = "modMain"
Public intZoom As Integer

Public intLocX As Integer
Public intLocY As Integer
Public intMapX As Integer
Public intMapY As Integer
Public intArrayPos As Integer
Public intCurrentArray As Integer

Public boolMapGrid As Boolean
Public boolSaved As Boolean
Public boolPainting As Boolean

Public strFilePath As String
Public strImagePath As String

Public intArrayGrid() As Integer
Public intMapGrid() As Integer
Public intArrayRGB() As Long

Public intMapClips As Integer
Public intImageHeight As Integer
Public intImageWidth As Integer

Sub Main()
    intZoom = 20
    intMapX = 0
    intMapY = 0
    intArrayPos = 0
    frmMain.Show
    boolMapGrid = False
    boolSaved = True
    strFilePath = Command()
End Sub
