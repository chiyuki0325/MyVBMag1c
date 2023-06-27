Attribute VB_Name = "modInit"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Sub Main()
    Dim Mag1c As New clsP1167
    Dim StartTick As Long, EndTick As Long
    MsgBox Mag1c.P1167(StartTick, EndTick)
    MsgBox "执行完成，消耗时间 " & (EndTick - StartTick) & " ms"
End Sub
