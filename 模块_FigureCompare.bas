Attribute VB_Name = "模块_FigureCompare"
Option Explicit

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                           绘制对比数据图                             ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/24

'更新内容：
'1.分离表名，如PKPM和YJK对比，将表明定为：F_P&Y

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2013/8/27

'更新内容：
'1.调用数组绘图程序


'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2013/8/24

'更新内容：
'1.增加迁移位移角数据代码，操作上不必通过各个程序绘图来生成；


Sub FigureCompare(model_1, model_2, Programe_1, Programe_2)

Dim fs As String

If model_1 = "d_P" And model_2 = "d_M" Then
    fs = "F_P&M"""
End If
If model_1 = "d_P" And model_2 = "d_Y" Then
    fs = "F_P&Y"""
End If
If model_1 = "d_Y" And model_2 = "d_M" Then
    fs = "F_Y&M"""
End If
If model_1 = "d_E" And model_2 = "d_M" Then
    fs = "F_E&M"""
End If
If model_1 = "d_P" And model_2 = "d_E" Then
    fs = "F_P&E"""
End If
If model_1 = "d_E" And model_2 = "d_Y" Then
    fs = "F_E&Y"""
End If


'If CheckBox4_PKPM And CheckBox4_MBuilding Then
'    Call FigureCompare("d_P", "d_M", "SATWE", "Midas Building")
'End If
'
'If CheckBox4_PKPM And CheckBox4_YJK Then
'    Call FigureCompare("d_P", "d_Y", "SATWE", "YJK")
'End If
'
'If CheckBox4_YJK And CheckBox4_MBuilding Then
'    Call FigureCompare("d_Y", "d_M", "YJK", "Midas Building")
'End If
'
'If CheckBox4_ETABS And CheckBox4_MBuilding Then
'    Call FigureCompare("d_E", "d_M", "ETABS", "Midas Building")
'End If
'
'If CheckBox4_ETABS And CheckBox4_PKPM Then
'    Call FigureCompare("d_P", "d_E", "SATWE", "ETABS")
'End If
'
'If CheckBox4_ETABS And CheckBox4_YJK Then
'    Call FigureCompare("d_E", "d_Y", "ETABS", "YJK")
'End If

'----------------------------------------------------------删除已有figure工作表
Dim sh As Worksheet

'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = fs Then
        sh.Delete
    End If
Next

'新建一个工作表，并命名为figure
With Worksheets
    Set sh = .Add(After:=Worksheets(.Count))
    sh.name = fs
    End With
'----------------------------------------------------------读取楼层总数
Num_all = Sheets(model_1).[A65536].End(xlUp).Row - 2


'----------------------------------------------------------迁移位移角数据
With Sheets(model_1).range("BI3:" & "BP" & Num_all + 2)
    .FormulaR1C1 = "=1/RC[-35]"
    .Font.ColorIndex = 1
    .Locked = True
End With
With Sheets(model_2).range("BI3:" & "BP" & Num_all + 2)
    .FormulaR1C1 = "=1/RC[-35]"
    .Font.ColorIndex = 1
    .Locked = True
End With

    
'----------------------------------------------------------对比绘图

'绘图公用部分
Dim sheetname(2), range_X(1), range_Y(0), name_S(1), name_XY(1)
Dim Location(1) As Integer

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284

'公用部分
sheetname(0) = model_1
sheetname(1) = model_2
sheetname(2) = fs
range_Y(0) = "R3C1:R" & Num_all + 2 & "C1"
name_S(0) = Programe_1
name_S(1) = Programe_2
name_XY(1) = "楼层"

'X向刚度比
range_X(0) = "R3C2:R" & Num_all + 2 & "C2"
range_X(1) = range_X(0)
name_XY(0) = "X向刚度比"
Location(0) = 0 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向刚度
range_X(0) = "R3C4:R" & Num_all + 2 & "C4"
range_X(1) = range_X(0)
name_XY(0) = "X向刚度"
Location(0) = 1 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下X向剪力(kN)
range_X(0) = "R3C6:R" & Num_all + 2 & "C6"
range_X(1) = range_X(0)
name_XY(0) = "风荷载下X向剪力(kN)"
Location(0) = 2 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下Y向剪力(kN)
range_X(0) = "R3C8:R" & Num_all + 2 & "C8"
range_X(1) = range_X(0)
name_XY(0) = "风荷载下Y向剪力(kN)"
Location(0) = 3 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下X向剪力(kN)
range_X(0) = "R3C10:R" & Num_all + 2 & "C10"
range_X(1) = range_X(0)
name_XY(0) = "地震作用下X向剪力(kN)"
Location(0) = 4 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下Y向剪力(kN)
range_X(0) = "R3C14:R" & Num_all + 2 & "C14"
range_X(1) = range_X(0)
name_XY(0) = "地震作用下Y向剪力(kN)"
Location(0) = 5 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向刚度比
range_X(0) = "R3C3:R" & Num_all + 2 & "C3"
range_X(1) = range_X(0)
name_XY(0) = "Y向刚度比"
Location(0) = 0 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向刚度
range_X(0) = "R3C5:R" & Num_all + 2 & "C5"
range_X(1) = range_X(0)
name_XY(0) = "Y向刚度"
Location(0) = 1 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下X向弯矩(kNm)
range_X(0) = "R3C7:R" & Num_all + 2 & "C7"
range_X(1) = range_X(0)
name_XY(0) = "风荷载下X向弯矩(kNm)"
Location(0) = 2 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下Y向弯矩(kNm)
range_X(0) = "R3C9:R" & Num_all + 2 & "C9"
range_X(1) = range_X(0)
name_XY(0) = "风荷载下Y向弯矩(kNm)"
Location(0) = 3 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下X向弯矩(kNm)
range_X(0) = "R3C11:R" & Num_all + 2 & "C11"
range_X(1) = range_X(0)
name_XY(0) = "地震作用下X向弯矩(kNm)"
Location(0) = 4 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下Y向弯矩(kNm)
range_X(0) = "R3C15:R" & Num_all + 2 & "C15"
range_X(1) = range_X(0)
name_XY(0) = "地震作用下Y向弯矩(kNm)"
Location(0) = 5 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX工况下位移角
range_X(0) = "R3C61:R" & Num_all + 2 & "C61"
range_X(1) = range_X(0)
name_XY(0) = "EX工况下位移角"
Location(0) = 0 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'EY工况下位移角
range_X(0) = "R3C65:R" & Num_all + 2 & "C65"
range_X(1) = range_X(0)
name_XY(0) = "EY工况下位移角"
Location(0) = 1 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'WX工况下位移角
range_X(0) = "R3C64:R" & Num_all + 2 & "C64"
range_X(1) = range_X(0)
name_XY(0) = "WX工况下位移角"
Location(0) = 2 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'WY工况下位移角
range_X(0) = "R3C68:R" & Num_all + 2 & "C68"
range_X(1) = range_X(0)
name_XY(0) = "WY工况下位移角"
Location(0) = 3 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'EX+工况下位移比
range_X(0) = "R3C35:R" & Num_all + 2 & "C35"
range_X(1) = range_X(0)
name_XY(0) = "EX+工况下位移比"
Location(0) = 0 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX-工况下位移比
range_X(0) = "R3C36:R" & Num_all + 2 & "C36"
range_X(1) = range_X(0)
name_XY(0) = "EX-工况下位移比"
Location(0) = 1 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY+工况下位移比
range_X(0) = "R3C38:R" & Num_all + 2 & "C38"
range_X(1) = range_X(0)
name_XY(0) = "EY+工况下位移比"
Location(0) = 2 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY-工况下位移比
range_X(0) = "R3C39:R" & Num_all + 2 & "C39"
range_X(1) = range_X(0)
name_XY(0) = "EY-工况下位移比"
Location(0) = 3 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX+工况下层间位移比
range_X(0) = "R3C41:R" & Num_all + 2 & "C41"
range_X(1) = range_X(0)
name_XY(0) = "EX+工况下层间位移比"
Location(0) = 0 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX-工况下层间位移比
range_X(0) = "R3C42:R" & Num_all + 2 & "C42"
range_X(1) = range_X(0)
name_XY(0) = "EX-工况下层间位移比"
Location(0) = 1 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY+工况下层间位移比
range_X(0) = "R3C44:R" & Num_all + 2 & "C44"
range_X(1) = range_X(0)
name_XY(0) = "EY+工况下层间位移比"
Location(0) = 2 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY-工况下层间位移比
range_X(0) = "R3C45:R" & Num_all + 2 & "C45"
range_X(1) = range_X(0)
name_XY(0) = "EY-工况下层间位移比"
Location(0) = 3 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向剪重比
range_X(0) = "R3C12:R" & Num_all + 2 & "C12"
range_X(1) = range_X(0)
name_XY(0) = "X向剪重比"
Location(0) = 0 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向剪重比
range_X(0) = "R3C16:R" & Num_all + 2 & "C16"
range_X(1) = range_X(0)
name_XY(0) = "Y向剪重比"
Location(0) = 1 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向抗剪承载力比
range_X(0) = "R3C46:R" & Num_all + 2 & "C46"
range_X(1) = range_X(0)
name_XY(0) = "X向抗剪承载力比"
Location(0) = 2 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向抗剪承载力比
range_X(0) = "R3C47:R" & Num_all + 2 & "C47"
range_X(1) = range_X(0)
name_XY(0) = "Y向抗剪承载力比"
Location(0) = 3 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'单位面积质量
range_X(0) = "R3C54:R" & Num_all + 2 & "C54"
range_X(1) = range_X(0)
name_XY(0) = "单位面积质量"
Location(0) = 4 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'质量比
range_X(0) = "R3C55:R" & Num_all + 2 & "C55"
range_X(1) = range_X(0)
name_XY(0) = "质量比"
Location(0) = 5 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱X向地震剪力百分比
range_X(0) = "R3C49:R" & Num_all + 2 & "C49"
range_X(1) = range_X(0)
name_XY(0) = "框架柱X向地震剪力百分比"
Location(0) = 0 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱Y向地震剪力百分比
range_X(0) = "R3C52:R" & Num_all + 2 & "C52"
range_X(1) = range_X(0)
name_XY(0) = "框架柱Y向地震剪力百分比"
Location(0) = 1 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱X向地震剪力调整系数
range_X(0) = "R3C50:R" & Num_all + 2 & "C50"
range_X(1) = range_X(0)
name_XY(0) = "框架柱X向地震剪力调整系数"
Location(0) = 2 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱Y向地震剪力调整系数
range_X(0) = "R3C53:R" & Num_all + 2 & "C53"
range_X(1) = range_X(0)
name_XY(0) = "框架柱Y向地震剪力调整系数"
Location(0) = 3 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)


End Sub

