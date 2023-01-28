Attribute VB_Name = "modLongestPrefix"
Sub Main()
    Dim Strs() As String
    Strs() = Split("Chi_Tang Chi_Zhao", " ")
    MsgBox GetLongestPrefix(Strs)
End Sub

Function GetLongestPrefix(ByRef Strs() As String) As String
    Dim Completed As Boolean, Prefix As String, Str As Variant, Count As Integer
    Completed = False
    Count = 1
    While Not Completed
        Prefix = Left(Strs(0), Count)
        For Each Str In Strs
            If Prefix <> Left(Str, Count) Then
                Completed = True
            End If
        Next
        Count = Count + 1
    Wend
    GetLongestPrefix = Left(Prefix, Len(Prefix) - 1)
End Function

