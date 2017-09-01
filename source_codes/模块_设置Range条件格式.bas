Attribute VB_Name = "模块_设置Range条件格式"
'设置Range条件格式的代码


'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2014/3/14

'更新内容：
'1.支持同时设置最大值与最小值的条件格式


'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2013/12/09

'更新内容：
'1.因为07版office出错，增加跳过错误语句

'变量说明:
'1.r为Range变量；
'2.maxormin输为"max"或者"min"；
'3.ser为 start and end string,max公式中的范围

Sub maxormin(R As range, maxormin As String, ser As String)

On Error Resume Next

Dim rr As String
rr = "=" & maxormin & "(" & ser & ")"

R.FormatConditions(R.FormatConditions.Count).SetFirstPriority

If maxormin = "max" Then

    With R.FormatConditions.Add(xlCellValue, xlEqual, rr)
        With R.FormatConditions(R.FormatConditions.Count)
            .Font.color = -16383844
            .Font.TintAndShade = 0
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.color = 13551615
            .Interior.TintAndShade = 0
        End With
    End With
    
    ElseIf maxormin = "min" Then
    
        With R.FormatConditions.Add(xlCellValue, xlEqual, rr)
            With R.FormatConditions(R.FormatConditions.Count)
                .Font.color = -16383844
                .Font.TintAndShade = 0
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.color = 10284031
                .Interior.TintAndShade = 0
            End With
            i = i + 1
        End With
End If

End Sub

