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
    QuestionCount = CInt(InputBox("��һ��һ������NN����ʾ���Online Judge����Ŀ����N��5000��"))
    If QuestionCount > 5000 Then
        Err.Raise 114514, "clsP1167.P1167", "�������࣬��Ҫû�ˣ�"
    End If
    
    Dim MinutePerQuestions%(), i%
    For i = 0 To QuestionCount - 1
    ReDim Preserve MinutePerQuestions(0 To i)
    MinutePerQuestions(i) = CInt(InputBox("֮������� " & i + 1 & " ����Ҫ�ķ�����"))
        If MinutePerQuestions(i) > 10000 Then
            Err.Raise 1919810, "clsP1167.P1167", "��Ŀ��ʱ��������������������?"
        End If
    Next
    
    Dim StartDateTime As Date, EndDateTime As Date
    StartDateTime = Me.DateTimeSerial(InputBox("�����뵱ǰʱ�䣨yyyy-mm-dd-hh:min��"))
    EndDateTime = Me.DateTimeSerial(InputBox("�����뾺��ʱ�䣨yyyy-mm-dd-hh:min��"))
    
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