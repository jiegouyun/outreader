Attribute VB_Name = "模块4"
Option Explicit

'从字符串sStr中提取第iNum个数字串的通用函数：
'修改自 officefans.net 『 VBA交流 』版主“小fisher”的代码

Function extractNumberFromString2(sStr As String, iNum As Integer) As Single
    Dim regEx '正则表达式对象（Regular Expression）
    Dim mc '匹配项集合（Match Collection）
    Set regEx = CreateObject("VBScript.RegExp") '建立新的正则表达式对象
    With regEx
        .Global = True '全局匹配，返回全部匹配项，如果为false只返回第一个匹配项
        .Pattern = "\s?(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)" '数字样式为[-]#[.#][E?]#
        Set mc = .Execute(sStr)  '执行匹配项查找
        '如果字符串中匹配项数目不小于iNum，则返回第iNum个匹配项
        If mc.Count >= iNum Then
            extractNumberFromString2 = mc(iNum - 1).Value
        '否则报告错误
        'Else
        '    Err.Raise 1, , "字符串中不存在相应数字"
        End If
    End With
End Function

