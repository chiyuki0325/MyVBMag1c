Attribute VB_Name = "modP1098"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Enum UnfoldModes
    LowerCase = 1
    UpperCase = 2
    Asterisk = 3
End Enum

Private Const chrCaseDelta As Byte = &H20
Private Const chrDashSymbol As Byte = &H2D  '-
Private Const chrAsterisk As Byte = &H2A  '*

Private Sub ArrayAppend(ByRef ArrayFrom() As Byte, CharToAppend As Byte)
    CopyMemory ByVal VarPtr(ArrayFrom(UBound(ArrayFrom))), CharToAppend, 1
End Sub

Sub P1098(ByRef StartTick As Long, ByRef EndTick As Long, ByRef OutputString As String)
'定义参数，输入参数并处理
    Dim UnfoldMode As UnfoldModes, RepeatLength As Integer, ReverseOrder As Boolean
    Dim InputArr() As String: InputArr = Split(InputBox("输入 p1, p2 和 p3"), " ")
    Dim strInput() As Byte: strInput = StrConv(InputBox("输入一个字符串"), vbFromUnicode)
    StartTick = GetTickCount
    Select Case InputArr(0)
    Case "1": UnfoldMode = LowerCase
    Case "2": UnfoldMode = UpperCase
    Case "3": UnfoldMode = Asterisk
    End Select
    RepeatLength = CInt(InputArr(1))
    Select Case InputArr(2)
    Case "1": ReverseOrder = False
    Case "2": ReverseOrder = True
    End Select
    'VB型字符串转为Byte()
    Dim strOutput() As Byte: ReDim strOutput(0 To 0)
    '遍历字符串并作出处理
    Dim i As Integer, j As Integer, chrTemp As Byte
    Dim chrFrom As Byte, chrTo As Byte, chrDestination As Byte, chrPlus As Integer
    Dim IsNumeric As Boolean, SyntaxError As Boolean
    For i = LBound(strInput) To UBound(strInput)
        If strInput(i) = chrDashSymbol Then
            chrFrom = strInput(i - 1)
            If i = UBound(strInput) Then
                SyntaxError = True
            Else
                chrTo = strInput(i + 1)
                If chrFrom >= &H30 And chrFrom <= &H39 Then  'From是数字
                    IsNumeric = True
                    If chrTo < &H30 Or chrTo > &H39 Then  'To不是数字
                        SyntaxError = True
                    End If
                Else
                    IsNumeric = False
                    If chrFrom >= &H61 And chrFrom <= &H7A Then  'From是小写字母
                        chrFrom = chrFrom - chrCaseDelta
                    ElseIf chrFrom < &H41 Or chrFrom > &H5A Then  'From不是大写字母
                        SyntaxError = True
                    End If
    
                    If chrTo >= &H61 And chrTo <= &H7A Then  'To是小写字母
                        chrTo = chrTo - chrCaseDelta
                    ElseIf chrTo < &H41 Or chrTo > &H5A Then  'To不是大写字母
                        SyntaxError = True
                    End If
                End If
            End If
            If chrTo <= chrFrom Then SyntaxError = True
            If Not SyntaxError Then
                If Not ReverseOrder Then
                    '正常顺序
                    chrTemp = chrFrom
                    chrDestination = (chrTo - 1)
                    chrPlus = 1
                Else
                    '逆序
                    chrTemp = chrTo
                    chrDestination = (chrFrom + 1)
                    chrPlus = -1
                End If
                Do Until chrTemp = chrDestination
                    chrTemp = chrTemp + chrPlus
                    For j = 1 To RepeatLength
                        If IsNumeric Then
                            If UnfoldMode = Asterisk Then
                                ArrayAppend strOutput, chrAsterisk
                            Else
                                ArrayAppend strOutput, chrTemp
                            End If
                        Else
                            Select Case UnfoldMode
                                Case UnfoldModes.LowerCase
                                    ArrayAppend strOutput, chrTemp + chrCaseDelta
                                Case UnfoldModes.UpperCase
                                    ArrayAppend strOutput, chrTemp
                                Case UnfoldModes.Asterisk
                                    ArrayAppend strOutput, chrAsterisk
                            End Select
                        End If
                        ReDim Preserve strOutput(LBound(strOutput) To UBound(strOutput) + 1)
                    Next
                Loop
            Else
                '小于等于，直接打一个减号
                ArrayAppend strOutput, chrDashSymbol
                ReDim Preserve strOutput(LBound(strOutput) To UBound(strOutput) + 1)
            End If
        Else
            '将这个字符加入到输出中
            ArrayAppend strOutput, strInput(i)
            ReDim Preserve strOutput(LBound(strOutput) To UBound(strOutput) + 1)
        End If
    Next
    DoEvents
    EndTick = GetTickCount
    'Byte()转VB型字符串
    OutputString = _
    "输入 1: " & Join(InputArr, ", ") & vbCrLf & _
    "输入 2: " & StrConv(strInput, vbUnicode) & vbCrLf & _
     "输出: " & StrConv(strOutput, vbUnicode)
End Sub
