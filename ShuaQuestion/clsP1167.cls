VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsP1167"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function P1167(ByRef StartTick, ByRef EndTick) As Integer
    Dim QuestionCount%
    QuestionCount = CInt(InputBox("第一行一个整数NN，表示洛谷Online Judge的题目数，N≤5000。"))
    If QuestionCount > 5000 Then
        Err.Raise 114514, "clsP1167.P1167", "题数过多，人要没了！"
    End If
    
    Dim MinutePerQuestions%(), i%
    For i = 0 To QuestionCount - 1
    ReDim Preserve MinutePerQuestions(0 To i)
    MinutePerQuestions(i) = CInt(InputBox("之后输入第 " & i + 1 & " 题需要的分钟数"))
        If MinutePerQuestions(i) > 10000 Then
            Err.Raise 1919810, "clsP1167.P1167", "题目耗时过长，这样还能做完吗?"
        End If
    Next
    
    Dim StartDateTime As Date, EndDateTime As Date
    StartDateTime = Me.DateTimeSerial(InputBox("请输入当前时间（yyyy-mm-dd-hh:min）"))
    EndDateTime = Me.DateTimeSerial(InputBox("请输入竞赛时间（yyyy-mm-dd-hh:min）"))
    
    StartTick = GetTickCount
    P1167 = Me.ShuaQuestion(QuestionCount, MinutePerQuestions, StartDateTime, EndDateTime)
    DoEvents
    EndTick = GetTickCount
End Function

Function ShuaQuestion( _
    QuestionCount%, _
    MinutePerQuestions%(), _
    StartDateTime As Date, _
    EndDateTime As Date _
) As Integer
    Dim ShuaingDateTime As Date
    Dim i%
    ShuaingDateTime = StartDateTime
    For i = 0 To QuestionCount - 1
        ShuaingDateTime = DateTime.DateAdd("n", MinutePerQuestions(i), ShuaingDateTime)
        If ShuaingDateTime > EndDateTime Then
            Exit For
        End If
        ShuaQuestion = ShuaQuestion + 1
    Next
End Function


Function DateTimeSerial(DateTimeString$) As Date
    '2007-06-23-11:59
    Dim DTArr$(), TimeArr$()
    DTArr = Split(DateTimeString, "-")
    TimeArr = Split(DTArr(3), ":")
    DateTimeSerial = DateTime.DateSerial(DTArr(0), DTArr(1), DTArr(2)) + DateTime.TimeSerial(TimeArr(0), TimeArr(1), 0)
End Function
