Attribute VB_Name = "modCommon"
Option Explicit

Sub DrawMap()
    
    If frmMain.WindowState = 1 Then Exit Sub
    
    Dim CounterX As Integer
    Dim CounterY As Integer
    Dim intGridSizeSource As Integer
    Dim dblGridSizeDestX As Double
    Dim dblGridSizeDestY As Double
    Dim lngGridSizeY As Long
    Dim lngGridSizeX As Long
    
    If frmMain.picclipArray.Picture > 0 Then
        
        If boolMapGrid = True Then
            lngGridSizeY = 1.5
            lngGridSizeX = 1.5
        Else
            lngGridSizeY = 0
            lngGridSizeX = 0
        End If
        
        intGridSizeSource = frmMain.picclipArray.Height
        dblGridSizeDestX = (frmMain.picGrid.Width - 2) / intZoom
        dblGridSizeDestY = (frmMain.picGrid.Height - 2) / intZoom
        If dblGridSizeDestX < 3 Then dblGridSizeDestX = 3
        If dblGridSizeDestY < 3 Then dblGridSizeDestY = 3
        If frmMain.picclipArray.Picture < 1 Then Exit Sub
        If intMapX < 1 Then Exit Sub
        If intMapY < 1 Then Exit Sub
        If intLocX < 0 Then intLocX = 0
        If intLocY < 0 Then intLocY = 0
        
        For CounterX = 0 To intZoom - 1
            For CounterY = 0 To intZoom - 1
                frmMain.picGrid.PaintPicture frmMain.picclipArray.Picture, _
                    CounterX * dblGridSizeDestX, _
                    CounterY * dblGridSizeDestY, _
                    Int(dblGridSizeDestX + 1) - lngGridSizeX, _
                    Int(dblGridSizeDestY + 1) - lngGridSizeY, _
                    intMapGrid(CounterX + intLocX, CounterY + intLocY) * frmMain.picclipArray.Height, _
                    0, _
                    frmMain.picclipArray.Height, _
                    frmMain.picclipArray.Height
            Next CounterY
        Next CounterX
        
    End If
    
    boolPainting = False
    
End Sub

Sub DrawArray()
    Dim intArrayHeight As Integer
    Dim Counter As Integer
    
    frmMain.vscrollArray.Min = 0
    frmMain.vscrollArray.Max = intMapClips - (frmMain.picArray.Height) / 33
    frmMain.vscrollArray.LargeChange = Int(frmMain.picArray.Height / 32)
    
    If frmMain.picclipArray.Picture > 0 Then
        intArrayHeight = Int(frmMain.picclipArray.Width / frmMain.picclipArray.Height)
        If frmMain.picArray.Height > intArrayHeight * 32 Then frmMain.picArray.Height = intArrayHeight * 32
        
        For Counter = 0 To Int(frmMain.picArray.Height / 33) - 1
            frmMain.picArray.PaintPicture frmMain.picclipArray.Picture, _
                0, Counter * 33, 32, 32, _
                (Counter + intArrayPos) * frmMain.picclipArray.Height, _
                0, _
                frmMain.picclipArray.Height, _
                frmMain.picclipArray.Height
        Next Counter
    End If
End Sub
