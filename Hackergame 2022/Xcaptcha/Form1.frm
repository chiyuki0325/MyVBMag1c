VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11250
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   4410
      TabIndex        =   2
      Top             =   7530
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   2340
      TabIndex        =   1
      Top             =   7470
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser ie 
      Height          =   6885
      Left            =   510
      TabIndex        =   0
      Top             =   480
      Width           =   10125
      ExtentX         =   17859
      ExtentY         =   12144
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()
    'Dim ie As New InternetExplorer, document As MSHTML.HTMLDocument
    Dim document As MSHTML.HTMLDocument
    'Sleep 500
    
    ie.navigate "http://202.38.93.111:10047/?token=" & MyToken, , , , "User-Agent: Mozilla/5.0 (X11; Linux x86_64; rv:108.0) Gecko/20100101 Firefox/108.0" '登录
    Do While ie.Busy And ie.readyState <> 4
    DoEvents
    Loop
    ie.navigate "http://202.38.93.111:10047/xcaptcha"
    Do While ie.Busy And ie.readyState <> 4
    DoEvents
    Loop
    
    'Debug.Print document.body.innerHTML
    
    Dim count As Integer
    For Each Label In ie.document.getElementsByTagName("label")
        count = count + 1
        Debug.Print Label.innerHTML
        expression = Split(Trim(Split(Label.innerHTML, " ")(0)), "+")
        ie.document.getElementById("captcha" & count).Value = BiggerAddition(expression(0), expression(1))
        Debug.Print "=" & ie.document.getElementById("captcha" & count).Value
    Next
    'Call ie.document.parentWindow.execScript("document.getElementById('submit').click();")
    'Sleep 500
    Do While ie.Busy And ie.readyState <> 4
    DoEvents
    Loop
    'MsgBox ie.document.body.innerHTML
        'ie.document.body.parentWindow.execScript "document.getElementById('submit').click(); alert('qwq');"
End Sub

'用字符串简单实现的大数相加

Function BiggerAddition(ByVal Num1 As String, ByVal Num2 As String) As String
    Dim Num1s() As Integer, Num2s() As Integer
    
    Dim i As Long
    
    For i = 0 To Len(Num1) - 1
        ReDim Preserve Num1s(0 To i)
        Num1s(i) = CInt(Left(Num1, 1))
        Num1 = Right(Num1, Len(Num1) - 1)
    Next
    
    For i = 0 To Len(Num2) - 1
        ReDim Preserve Num2s(0 To i)
        Num2s(i) = CInt(Left(Num2, 1))
        Num2 = Right(Num2, Len(Num2) - 1)
    Next
    
    If UBound(Num1s) > UBound(Num2s) Then
        ReDim Preserve Num2s(0 To UBound(Num1s))
    Else
        ReDim Preserve Num1s(0 To UBound(Num2s))
    End If
    
    Dim Subscript As Long, qwq As String, tmps As String, tmp2 As String
    tmp2 = "0"
    For i = 0 To UBound(Num1s)
        Subscript = UBound(Num1s) - i
        tmps = CStr(Num1s(Subscript) + Num2s(Subscript) + CLng(tmp2))
        Debug.Print tmps
        If Len(tmps) > 1 Then
            tmp2 = Left(tmps, Len(tmps) - 1)
        Else
            tmp2 = "0"
        End If
        qwq = Right(tmps, 1) & qwq
    Next
    
    BiggerAddition = qwq
End Function


Private Sub Command1_Click()
Main
End Sub

Private Sub Command2_Click()

    Dim count As Integer
    For Each Label In ie.document.getElementsByTagName("label")
        count = count + 1
        Debug.Print Label.innerHTML
        expression = Split(Trim(Split(Label.innerHTML, " ")(0)), "+")
        ie.document.getElementById("captcha" & count).Value = BiggerAddition(expression(0), expression(1))
        Debug.Print "=" & ie.document.getElementById("captcha" & count).Value
    Next
End Sub
