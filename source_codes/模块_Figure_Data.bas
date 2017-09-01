Attribute VB_Name = "模块_Figure_Data"
Option Explicit

'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            分布数据绘图                              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/03
'1.迁移处理位移角数据中，.FormulaR1C1 = "=1/d_P!RC[-35]"改为.FormulaR1C1 = "=1/" & dis_sheet & "!RC[-35]"


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27
'1.更改为通用绘图函数，传递时程数据软件名可以画相应软件下的数据曲线

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/13
'1.修改位移角绘图；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/2
'1.增加绘图

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/27/ 21:57
'更新内容:
'1.定义Num_All从d_P中第一列读取最大楼层数



Sub OUTReader_Figure_Data(softname)

'----------------------------------------------------------删除已有figure工作表
Dim sh As Worksheet

Dim dis_sheet, fig_sheet As String

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284

If softname = "PKPM" Then
    dis_sheet = "d_P"
    fig_sheet = "figure_PKPM"
ElseIf softname = "YJK" Then
    dis_sheet = "d_Y"
    fig_sheet = "figure_YJK"
ElseIf softname = "MBuilding" Then
    dis_sheet = "d_M"
    fig_sheet = "figure_MBuilding"
ElseIf softname = "ETABS" Then
    dis_sheet = "d_E"
    fig_sheet = "figure_ETABS"
Else
    MsgBox "参数输入错误"
End If

'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = fig_sheet Then
        sh.Delete
    End If
Next

'新建一个工作表，并命名为figure
With Worksheets
    Set sh = .Add(After:=Worksheets(.Count))
    sh.name = fig_sheet
    End With

'----------------------------------------------------------迁移处理位移角数据
'Sheets("figure_PKPM").range("A3:" & "A" & Num_All + 2).Value = Sheets("d_P").range("A3:" & "A" & Num_All + 2).Value
'Sheets("figure_PKPM").range("A3:" & "A" & Num_All + 2).Font.ColorIndex = 2
'Sheets("figure_PKPM").range("A3:" & "A" & Num_All + 2).Locked = True
With Sheets(dis_sheet).range("BI3:" & "BP" & Num_all + 2)
    .FormulaR1C1 = "=1/" & dis_sheet & "!RC[-35]"
    .Font.ColorIndex = 1
    .Locked = True
End With

'----------------------------------------------------------读取楼层总数
Num_all = Sheets(dis_sheet).[A65536].End(xlUp).Row - 2


'----------------------------------------------------------调用过程绘图
'X向刚度比、刚度及层间剪力
Call add_chart(softname, "B3:" & "B" & Num_all + 2, "A3:" & "A" & Num_all + 2, "X向刚度比", "刚度比", "层号", 0 * Width, 0 * Hight, Width, Hight)
Call add_chart(softname, "D3:" & "D" & Num_all + 2, "A3:" & "A" & Num_all + 2, "X向刚度", "刚度", "层号", 1 * Width, 0 * Hight, Width, Hight)
Call add_chart(softname, "F3:" & "F" & Num_all + 2, "A3:" & "A" & Num_all + 2, "风荷载下X向剪力", "剪力(kN)", "层号", 2 * Width, 0 * Hight, Width, Hight)
Call add_chart(softname, "H3:" & "H" & Num_all + 2, "A3:" & "A" & Num_all + 2, "风荷载下Y向剪力", "剪力(kN)", "层号", 3 * Width, 0 * Hight, Width, Hight)
Call add_chart(softname, "J3:" & "J" & Num_all + 2, "A3:" & "A" & Num_all + 2, "地震作用下X向剪力", "剪力(kN)", "层号", 4 * Width, 0 * Hight, Width, Hight)
Call add_chart(softname, "N3:" & "N" & Num_all + 2, "A3:" & "A" & Num_all + 2, "地震作用下Y向剪力", "剪力(kN)", "层号", 5 * Width, 0 * Hight, Width, Hight)

'Y向刚度比、刚度及层间弯矩
Call add_chart(softname, "C3:" & "C" & Num_all + 2, "A3:" & "A" & Num_all + 2, "Y向刚度比", "刚度比", "层号", 0 * Width, 1 * Hight, Width, Hight)
Call add_chart(softname, "E3:" & "E" & Num_all + 2, "A3:" & "A" & Num_all + 2, "Y向刚度", "刚度", "层号", 1 * Width, 1 * Hight, Width, Hight)
Call add_chart(softname, "G3:" & "G" & Num_all + 2, "A3:" & "A" & Num_all + 2, "风荷载下X向弯矩", "弯矩(kNm)", "层号", 2 * Width, 1 * Hight, Width, Hight)
Call add_chart(softname, "I3:" & "I" & Num_all + 2, "A3:" & "A" & Num_all + 2, "风荷载下Y向弯矩", "弯矩(kNm)", "层号", 3 * Width, 1 * Hight, Width, Hight)
Call add_chart(softname, "K3:" & "K" & Num_all + 2, "A3:" & "A" & Num_all + 2, "地震作用下X向弯矩", "弯矩(kNm)", "层号", 4 * Width, 1 * Hight, Width, Hight)
Call add_chart(softname, "O3:" & "O" & Num_all + 2, "A3:" & "A" & Num_all + 2, "地震作用下Y向弯矩", "弯矩(kNm)", "层号", 5 * Width, 1 * Hight, Width, Hight)

'层间位移角
Call add_chart(softname, "BI3:" & "BI" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EX工况下位移角", "位移角", "层号", 0 * Width, 2 * Hight, Width, Hight, "#/###0")
Call add_chart(softname, "BM3:" & "BM" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EY工况下位移角", "位移角", "层号", 1 * Width, 2 * Hight, Width, Hight, "#/###0")
Call add_chart(softname, "BL3:" & "BL" & Num_all + 2, "A3:" & "A" & Num_all + 2, "WX工况下位移角", "位移角", "层号", 2 * Width, 2 * Hight, Width, Hight, "#/###0")
Call add_chart(softname, "BP3:" & "BP" & Num_all + 2, "A3:" & "A" & Num_all + 2, "WY工况下位移角", "位移角", "层号", 3 * Width, 2 * Hight, Width, Hight, "#/###0")

'位移比
Call add_chart(softname, "AI3:" & "AI" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EX+工况下位移比", "位移比", "层号", 0 * Width, 3 * Hight, Width, Hight)
Call add_chart(softname, "AJ3:" & "AJ" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EX-工况下位移比", "位移比", "层号", 1 * Width, 3 * Hight, Width, Hight)
Call add_chart(softname, "AL3:" & "AL" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EY+工况下位移比", "位移比", "层号", 2 * Width, 3 * Hight, Width, Hight)
Call add_chart(softname, "AM3:" & "AM" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EY-工况下位移比", "位移比", "层号", 3 * Width, 3 * Hight, Width, Hight)

'层间位移比
Call add_chart(softname, "AO3:" & "AO" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EX+工况下层间位移比", "位移比", "层号", 0 * Width, 4 * Hight, Width, Hight)
Call add_chart(softname, "AP3:" & "AP" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EX-工况下层间位移比", "位移比", "层号", 1 * Width, 4 * Hight, Width, Hight)
Call add_chart(softname, "AR3:" & "AR" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EY+工况下层间位移比", "位移比", "层号", 2 * Width, 4 * Hight, Width, Hight)
Call add_chart(softname, "AS3:" & "AS" & Num_all + 2, "A3:" & "A" & Num_all + 2, "EY-工况下层间位移比", "位移比", "层号", 3 * Width, 4 * Hight, Width, Hight)

'剪重比、抗剪承载力比及质量比
Call add_chart(softname, "L3:" & "L" & Num_all + 2, "A3:" & "A" & Num_all + 2, "X向剪重比", "剪重比", "层号", 0 * Width, 5 * Hight, Width, Hight)
Call add_chart(softname, "P3:" & "P" & Num_all + 2, "A3:" & "A" & Num_all + 2, "Y向剪重比", "剪重比", "层号", 1 * Width, 5 * Hight, Width, Hight)
Call add_chart(softname, "AT3:" & "AT" & Num_all + 2, "A3:" & "A" & Num_all + 2, "X向抗剪承载力比", "抗剪承载力比", "层号", 2 * Width, 5 * Hight, Width, Hight)
Call add_chart(softname, "AU3:" & "AU" & Num_all + 2, "A3:" & "A" & Num_all + 2, "Y向抗剪承载力比", "抗剪承载力比", "层号", 3 * Width, 5 * Hight, Width, Hight)
Call add_chart(softname, "BB3:" & "BB" & Num_all + 2, "A3:" & "A" & Num_all + 2, "单位面积质量", "单位面积质量", "层号", 4 * Width, 5 * Hight, Width, Hight)
Call add_chart(softname, "BC3:" & "BC" & Num_all + 2, "A3:" & "A" & Num_all + 2, "质量比", "质量比", "层号", 5 * Width, 5 * Hight, Width, Hight)

'框架剪力所占总剪力比例及调整系数
Call add_chart(softname, "AW3:" & "AW" & Num_all + 2, "A3:" & "A" & Num_all + 2, "框架柱X向地震剪力百分比", "框架柱剪力百分比", "层号", 0 * Width, 6 * Hight, Width, Hight)
Call add_chart(softname, "AZ3:" & "AZ" & Num_all + 2, "A3:" & "A" & Num_all + 2, "框架柱Y向地震剪力百分比", "框架柱剪力百分比", "层号", 1 * Width, 6 * Hight, Width, Hight)
Call add_chart(softname, "AX3:" & "AX" & Num_all + 2, "A3:" & "A" & Num_all + 2, "框架柱X向地震剪力调整系数", "框架柱剪力调整系数", "层号", 2 * Width, 6 * Hight, Width, Hight)
Call add_chart(softname, "BA3:" & "BA" & Num_all + 2, "A3:" & "A" & Num_all + 2, "框架柱Y向地震剪力调整系数", "框架柱剪力调整系数", "层号", 3 * Width, 6 * Hight, Width, Hight)

End Sub
