Attribute VB_Name = "模块_CheckRegExpfromString"
Option Explicit

'判断字符串sStr中是否含有正则表达式Reg代表的模糊内容的通用函数，返回True或False

Function CheckRegExpfromString(sStr As String, Reg As String) As Boolean
    CheckRegExpfromString = False
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = Reg
        If .TEST(sStr) = True Then
            CheckRegExpfromString = True
        End If
    End With
End Function
