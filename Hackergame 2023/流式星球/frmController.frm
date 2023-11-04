VERSION 5.00
Begin VB.Form frmController 
   BackColor       =   &H80000005&
   Caption         =   "流式星人的视频控制器"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "播放"
      Height          =   420
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "初始化"
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1110
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Text            =   "759"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Text            =   "427"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "播放一帧"
      Height          =   420
      Left            =   1380
      TabIndex        =   1
      Top             =   2040
      Width           =   1170
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Text            =   "C:\Users\yidaozhan\Desktop\hackergame.bin"
      Top             =   120
      Width           =   3195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帧编号"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分辨率"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   780
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "视频路径"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   960
   End
   Begin VB.Label lblFrame 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   1380
      Width           =   105
   End
End
Attribute VB_Name = "frmController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TriggerSingleFrame()
On Error GoTo PlayCompleted
    lblFrame = FrameCount
    Form1.DrawCurrentFrame
    FrameCount = FrameCount + 1
    Exit Sub
PlayCompleted:
    MsgBox "播放完成"
End Sub

Private Sub Command1_Click()
    TriggerSingleFrame
End Sub

Private Sub Command2_Click()
    Form1.Refresh
    FrameCount = 0
    lblFrame = 0
    VideoWidth = CLng(txtWidth)
    VideoHeight = CLng(txtHeight)
    FrameSize = VideoWidth * VideoHeight * 3
    Form1.Width = VideoWidth * Screen.TwipsPerPixelX + 500
    Form1.Height = VideoHeight * Screen.TwipsPerPixelY + 500
    Open txtPath For Binary Access Read As #1
        ReDim Buffer(LOF(1))
        Get #1, , Buffer
    Close #1
End Sub

Private Sub Command3_Click()
    While True
        TriggerSingleFrame
    Wend
End Sub

