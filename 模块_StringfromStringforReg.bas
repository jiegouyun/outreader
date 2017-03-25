Attribute VB_Name = "模块_StringfromStringforReg"
Option Explicit

'从字符串sStr中按正则表达式Reg提取第iNum个字符串的通用函数

Function StringfromStringforReg(sStr As String, Reg As String, iNum As Integer) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = Reg
        Dim mc
        Set mc = .Execute(sStr)  '执行匹配项查找
        If mc.Count >= iNum Then
            StringfromStringforReg = mc(iNum - 1).Value
        End If
    End With
End Function
