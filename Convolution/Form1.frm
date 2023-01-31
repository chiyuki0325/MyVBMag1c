VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "赛博包浆制造机"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4185
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "woc, 卷"
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   1680
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   150
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1125
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "6"
      Height          =   135
      Left            =   2640
      TabIndex        =   3
      Top             =   210
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   135
      Left            =   1350
      TabIndex        =   2
      Top             =   210
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim X As Long, Y As Long, StartX As Long, StartY As Long
StartX = Pixelize(Label1.Left)
StartY = Pixelize(Label1.Top)

For X = 0 To Pixelize(Picture1.Width)
    
    For Y = 0 To Pixelize(Picture1.Height)
        Me.PSet (Twipize(StartX + X), Twipize(StartY + Y)), ConvolveSinglePixelGaussian(X, Y, Picture1)
        DoEvents
    Next

Next


StartX = Pixelize(Label2.Left)
StartY = Pixelize(Label2.Top)

For X = 0 To Pixelize(Picture1.Width)
    
    For Y = 0 To Pixelize(Picture1.Height)
        Me.PSet (Twipize(StartX + X), Twipize(StartY + Y)), ConvolveSinglePixelBorder(X, Y, Picture1)
        DoEvents
    Next

Next
End Sub


Function ConvolveSinglePixelGaussian(X As Long, Y As Long, ByRef TargetPicture As PictureBox) As Single
    Dim Reds(0 To 24) As Integer, Greens(0 To 24) As Integer, Blues(0 To 24) As Integer
    ' 0  1  2  3  4
    ' 5  6  7  8  9
    ' 10 11 12 13 14
    ' 15 16 17 18 19
    ' 20 21 22 23 24
    Dim Subscript As Integer: Subscript = 0
    Dim i As Integer, j As Integer
    '从图中取这 25 个格的像素以便继续操作
    For j = -2 To 2
        For i = -2 To 2
            GetSinglePixel Reds, Greens, Blues, Subscript, Twipize(X), Twipize(Y), Twipize(i), Twipize(j), TargetPicture
            Subscript = Subscript + 1
        Next
    Next
    
    Dim FinalRed As Double, FinalGreen As Double, FinalBlue As Double
    FinalRed = ConvolveSingleColorGaussian(Reds)
    FinalGreen = ConvolveSingleColorGaussian(Greens)
    FinalBlue = ConvolveSingleColorGaussian(Blues)
   
    'Debug.Print ("R" & FinalRed & " G" & FinalGreen & " B" & FinalBlue)
    
    ConvolveSinglePixelGaussian = RGB(Int(FinalRed), Int(FinalGreen), Int(FinalBlue))
    
End Function

Function ConvolveSinglePixelBorder(X As Long, Y As Long, ByRef TargetPicture As PictureBox) As Single
    Dim Reds(0 To 24) As Integer, Greens(0 To 24) As Integer, Blues(0 To 24) As Integer
    ' 0  1  2  3  4
    ' 5  6  7  8  9
    ' 10 11 12 13 14
    ' 15 16 17 18 19
    ' 20 21 22 23 24
    Dim Subscript As Integer: Subscript = 0
    Dim i As Integer, j As Integer
    '从图中取这 25 个格的像素以便继续操作
    For j = -2 To 2
        For i = -2 To 2
            GetSinglePixel Reds, Greens, Blues, Subscript, Twipize(X), Twipize(Y), Twipize(i), Twipize(j), TargetPicture
            Subscript = Subscript + 1
        Next
    Next
    
    Dim FinalRed As Double, FinalGreen As Double, FinalBlue As Double
    FinalRed = ConvolveSingleColorBorder(Reds)
    FinalGreen = ConvolveSingleColorBorder(Greens)
    FinalBlue = ConvolveSingleColorBorder(Blues)
   
    'Debug.Print ("R" & FinalRed & " G" & FinalGreen & " B" & FinalBlue)
    
    ConvolveSinglePixelBorder = RGB(Int(FinalRed), Int(FinalGreen), Int(FinalBlue))
    
End Function

Function Pixelize(Twip) As Integer
    Pixelize = Twip / Screen.TwipsPerPixelX
End Function

Function Twipize(Pixel) As Long
    Twipize = Pixel * Screen.TwipsPerPixelX
End Function

Sub GetSinglePixel(ByRef Reds() As Integer, ByRef Greens() As Integer, ByRef Blues() As Integer, Subscript As Integer, X As Long, Y As Long, i As Integer, j As Integer, ByRef TargetPicture As PictureBox)
    Dim Colors As Long
    Colors = TargetPicture.Point(X + i, Y + j)
    Reds(Subscript) = Colors And RGB(255, 0, 0)
    Greens(Subscript) = Int((Colors And RGB(0, 255, 0)) / 256)
    Blues(Subscript) = Int(Int((Colors And RGB(0, 0, 255)) / 256) / 256)
    'Debug.Print (X + i & " " & Y + j) & (" R" & Reds(Subscript) & " G" & Greens(Subscript) & " B" & Blues(Subscript))
End Sub

Function ConvolveSingleColorGaussian(ByRef Colors() As Integer) As Double
    ConvolveSingleColorGaussian = 0.003 * Colors(0) + 0.013 * Colors(1) + 0.022 * Colors(2) + 0.013 * Colors(3) + 0.003 * Colors(4) _
                        + 0.013 * Colors(5) + 0.06 * Colors(6) + 0.098 * Colors(7) + 0.06 * Colors(8) + 0.013 * Colors(9) _
                        + 0.022 * Colors(10) + 0.098 * Colors(11) + 0.162 * Colors(12) + 0.098 * Colors(13) + 0.022 * Colors(14) _
                        + 0.013 * Colors(15) + 0.06 * Colors(16) + 0.098 * Colors(17) + 0.06 * Colors(18) + 0.013 * Colors(19) _
                        + 0.003 * Colors(20) + 0.013 * Colors(21) + 0.022 * Colors(22) + 0.013 * Colors(23) + 0.003 * Colors(24)
End Function

Function ConvolveSingleColorBorder(ByRef Colors() As Integer) As Double
    ConvolveSingleColorBorder = 0 + 0 - 0.2 * Colors(2) + 0 + 0 _
                        + 0 - 0.2 * Colors(6) - 0.5 * Colors(7) - 0.2 * Colors(8) + 0 _
                        - 0.2 * Colors(10) - 0.5 * Colors(11) + 5 * Colors(12) - 0.5 * Colors(13) - 0.2 * Colors(14) _
                        + 0 - 0.2 * Colors(16) - 0.5 * Colors(17) - 0.2 * Colors(18) + 0 _
                        + 0 + 0 - 0.2 * Colors(22) + 0 + 0
                        ConvolveSingleColorBorder = 0.7 * Abs(ConvolveSingleColorBorder)
End Function

Private Sub Form_Load()
Label1.Caption = ""
Label2.Caption = ""
End Sub

Private Sub Label3_Click()

End Sub
