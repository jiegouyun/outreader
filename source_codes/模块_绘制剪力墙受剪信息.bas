Attribute VB_Name = "模块_绘制剪力墙受剪信息"
Option Explicit


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                          绘制剪力墙受剪信息代码                      ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/3
'1.分成两个模块，按楼层和按构件
'2.表格名更新


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/19
'更新内容:
'1.NN的定义由原来的A列改为B列
'2.图形名改为“抗剪截面要求”及“抗剪承载力”


Sub member_wall_f(softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

'sheetname() 为曲线数据所在sheet名，1 X N_C+1 维数组，最后一个值为图表所在sheet名
'N_C 为图中曲线总数
'range_X() 为图表的X轴数据range，1 X N_C 维数组
'range_Y() 为图表的Y轴数据range，1 X 1 维数组
'name() 为数据Series的标题，1 X N_C 维数组
'name_XY() 1 X 2 维数组，第一个值为图表的X轴的标题，第二个值为Y轴标题
'Location() 1 X 2 维数组，为图标在sheet中的位置；
'Optional NumFormat As String = "G/通用格式" 为X轴数据的格式，缺省值为通用，可为分数及科学计数等；


Dim NN As Integer: NN = Sheets(wallshearsheet).range("b65536").End(xlUp)

'绘图公用部分
Dim sheetname(), range_X(), range_Y(0), name_S(), name_XY(1)
ReDim Preserve sheetname(3)
ReDim Preserve range_X(2)
ReDim Preserve name_S(2)
Dim Location(1) As Integer

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284

Dim i As Integer

For i = 0 To 2
    sheetname(i) = wallshearsheet
Next i
sheetname(3) = wallshearsheet

range_Y(0) = "R4C2:R" & NN + 3 & "C2"
name_XY(1) = "层数"

       
range_X(0) = "R4C11:R" & NN + 3 & "C11"
name_S(0) = "剪力"

range_X(1) = "R4C24:R" & NN + 3 & "C24"
name_S(1) = "抗剪承载力"

range_X(2) = "R4C25:R" & NN + 3 & "C25"
name_S(2) = "抗剪截面要求"
    
                
name_XY(0) = "剪力墙受剪验算"

Location(0) = 1400
Location(1) = 100

Call add_chart_array(sheetname(), 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

End Sub


Sub member_wall_m(softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

'sheetname() 为曲线数据所在sheet名，1 X N_C+1 维数组，最后一个值为图表所在sheet名
'N_C 为图中曲线总数
'range_X() 为图表的X轴数据range，1 X N_C 维数组
'range_Y() 为图表的Y轴数据range，1 X 1 维数组
'name() 为数据Series的标题，1 X N_C 维数组
'name_XY() 1 X 2 维数组，第一个值为图表的X轴的标题，第二个值为Y轴标题
'Location() 1 X 2 维数组，为图标在sheet中的位置；
'Optional NumFormat As String = "G/通用格式" 为X轴数据的格式，缺省值为通用，可为分数及科学计数等；


Dim NN As Integer: NN = Sheets(wallshearsheet).range("c65536").End(xlUp)

'绘图公用部分
Dim sheetname(), range_X(), range_Y(0), name_S(), name_XY(1)
ReDim Preserve sheetname(3)
ReDim Preserve range_X(2)
ReDim Preserve name_S(2)
Dim Location(1) As Integer

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284

Dim i As Integer

For i = 0 To 2
    sheetname(i) = wallshearsheet
Next i
sheetname(3) = wallshearsheet

range_Y(0) = "R4C3:R" & NN + 3 & "C3"
name_XY(1) = "层数"

       
range_X(0) = "R4C11:R" & NN + 3 & "C11"
name_S(0) = "剪力"

range_X(1) = "R4C24:R" & NN + 3 & "C24"
name_S(1) = "抗剪承载力"

range_X(2) = "R4C25:R" & NN + 3 & "C25"
name_S(2) = "抗剪截面要求"
    
                
name_XY(0) = "剪力墙受剪验算"

Location(0) = 1400
Location(1) = 100

Call add_chart_array(sheetname(), 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

End Sub
