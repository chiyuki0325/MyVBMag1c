VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "流式星人的视频播放器"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmController.Show
    frmController.SetFocus
End Sub

Sub DrawCurrentFrame()
    Me.Refresh
    Dim i As Long, idx As Long
    Dim ThisLinePixels As Long, LineNum As Long
    LineNum = 1
    ThisLinePixels = 0
    For i = 1 To FrameSize Step 3
        idx = (FrameSize * FrameCount) + i - 1
        Me.PSet (ThisLinePixels, LineNum), _
                RGB(Buffer(idx + 2), Buffer(idx + 1), Buffer(idx))
        If ThisLinePixels = VideoWidth - 1 Then
            ThisLinePixels = 0
            LineNum = LineNum + 1
        Else
            ThisLinePixels = ThisLinePixels + 1
        End If
    Next
    DoEvents
End Sub
