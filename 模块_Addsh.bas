Attribute VB_Name = "模块_Addsh"
Option Explicit
Sub Addsh(name)

'定义工作表
Dim sh As Worksheet

'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = name Then
        Exit Sub
    End If
Next

'新建一个工作表，并命名为name
With Worksheets
    Set sh = .Add(After:=Worksheets(.Count))
    sh.name = name
    End With
    
End Sub

'调用新增工作表函数

Sub testforaddsh()

'括号内为新增工作表名称
Call Addsh("general")
Call Addsh("distribution1")

End Sub
