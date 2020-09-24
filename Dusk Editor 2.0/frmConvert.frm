VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Convert Values"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Convert Values"
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdNewDown 
         Caption         =   "<"
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdOldDown 
         Caption         =   "<"
         Height          =   495
         Left            =   3360
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdNewUp 
         Caption         =   ">"
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdOldUp 
         Caption         =   ">"
         Height          =   495
         Left            =   4080
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox picNew 
         Height          =   495
         Left            =   3600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox picOld 
         Height          =   495
         Left            =   3600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtOld 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblNewImage 
         Caption         =   "New Image"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblOldImage 
         Caption         =   "Original Image"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblOldValue 
         Caption         =   "Original Value"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNew 
         Caption         =   "New Value"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmConvert
End Sub

Private Sub cmdConvert_Click()
    Dim XCounter As Integer
    Dim YCounter As Integer
    For XCounter = 0 To intMapX
        For YCounter = 0 To intMapY
            If intMapGrid(XCounter, YCounter) = Val(txtOld.Text) Then
                intMapGrid(XCounter, YCounter) = Val(txtNew.Text)
            End If
        Next YCounter
    Next XCounter
    Unload frmConvert
End Sub

Private Sub cmdNewDown_Click()
    txtNew.Text = Val(txtNew.Text) - 1
End Sub

Private Sub cmdNewUp_Click()
    txtNew.Text = Val(txtNew.Text) + 1
End Sub

Private Sub cmdOldDown_Click()
    txtOld.Text = Val(txtOld.Text) - 1
End Sub

Private Sub cmdOldUp_Click()
    txtOld.Text = Val(txtOld.Text) + 1
End Sub

Private Sub Form_Load()
    Dim intValue As Integer
    txtOld.Text = intCurrentArray
    txtNew.Text = txtOld.Text
    intValue = Val(txtOld.Text)
    picOld.PaintPicture frmMain.picclipArray.Picture, _
        0, 0, 32, 32, _
        intValue * frmMain.picclipArray.Height, _
        0, _
        frmMain.picclipArray.Height, _
        frmMain.picclipArray.Height
End Sub

Private Sub txtNew_Change()
    Dim intValue As Integer
    If Val(txtNew.Text) < 0 Then txtNew.Text = 0
    If Val(txtNew.Text) > frmMain.picclipArray.Width / frmMain.picclipArray.Height - 1 Then
        txtNew.Text = Int(frmMain.picclipArray.Width / frmMain.picclipArray.Height) - 1
    End If
    intValue = Val(txtNew.Text)
    picNew.PaintPicture frmMain.picclipArray.Picture, _
        0, 0, 32, 32, _
        intValue * frmMain.picclipArray.Height, _
        0, _
        frmMain.picclipArray.Height, _
        frmMain.picclipArray.Height
End Sub

Private Sub txtOld_Change()
    Dim intValue As Integer
    If Val(txtOld.Text) < 0 Then txtOld.Text = 0
    If Val(txtOld.Text) > frmMain.picclipArray.Width / frmMain.picclipArray.Height - 1 Then
        txtOld.Text = Int(frmMain.picclipArray.Width / frmMain.picclipArray.Height) - 1
    End If
    intValue = Val(txtOld.Text)
    picOld.PaintPicture frmMain.picclipArray.Picture, _
        0, 0, 32, 32, _
        intValue * frmMain.picclipArray.Height, _
        0, _
        frmMain.picclipArray.Height, _
        frmMain.picclipArray.Height
End Sub
