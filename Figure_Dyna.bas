Attribute VB_Name = "Figure_Dyna"

Option Explicit


'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            时程分析绘图                              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////
'更新时间:2014/1/8
'1.定义全局变量楼层总数，确保在没有反应谱数据时时程画图楼层正确

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/11/29
'1.添加正负35%，正负20%反应谱剪力曲线

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/28
'1.修正没有反应数据时画图中还有反应谱系列名称bug

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27
'1.更改为通用绘图函数，传递时程数据sheet名可以画此图曲线

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/25
'更新内容:
'1.调整了位移角迁移代码；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/16/ 19:23
'更新内容:
'1.迁移来的位移角数据颜色设为白色,变相藏一下;



Sub OUTReader_Figure_Dyna(ela)

'----------------------------------------------------------删除已有figure工作表
Dim sh As Worksheet

'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = "figure_dyna" Then
        sh.Delete
    End If
Next

'新建一个工作表，并命名为figure
With Worksheets
    Set sh = .Add(After:=Worksheets(.Count))
    sh.name = "figure_dyna"
    End With

    
'确定楼层数
Sheets(ela).Select
Dim NN As Integer: NN = Cells(Rows.Count, "j").End(3).Row - 2
'Debug.Print "NN=" & NN
'定义全局变量楼层总数，确保在没有反应谱数据时时程画图楼层正确
Num_all = NN
'----------------------------------------------------------迁移处理位移角数据
Dim i As Integer
Dim num_th, num_th1 As Integer
num_th = Sheets(ela).Cells(2, 2)
num_th1 = num_th
Debug.Print num_th

'如没有反应谱数据，绘制曲线数减一
If Sheets(ela).Cells(num_th + 8, 2) = "" Then
    num_th = num_th - 1
End If

For i = 0 To 2 * (num_th + 3) - 1
    Sheets(ela).range(Sheets(ela).Cells(3, 100 + i), Sheets(ela).Cells(NN + 2, 100 + i)).Value = _
    Sheets(ela).range(Sheets(ela).Cells(3, 10 + 3 * i), Sheets(ela).Cells(NN + 2, 10 + 3 * i)).Value
Next

With Sheets(ela).range(Sheets(ela).Cells(3, 100 + 2 * (num_th + 3) + 1), Sheets(ela).Cells(NN + 2, 100 + 4 * (num_th + 3)))
    .FormulaR1C1 = "=1/RC[-" & 2 * (num_th + 3) + 1 & "]"
    .Font.ColorIndex = 2
    .Locked = True
End With

With Sheets(ela).range(Sheets(ela).Cells(3, 100), Sheets(ela).Cells(NN + 2, 100 + 4 * (num_th + 3)))
    .Font.ColorIndex = 2
    .Locked = True
End With


Sheets("figure_dyna").Select

'绘图公用部分
Dim sheetname(), range_X(), range_Y(0), name_S(), name_XY(1)
ReDim Preserve sheetname(num_th + 7)
ReDim Preserve range_X(num_th + 6)
ReDim Preserve name_S(num_th + 6)
Dim Location(1) As Integer

'图表高宽
Dim Width As Integer, Hight As Integer

Width = 414
Hight = 510

For i = 0 To num_th + 6
    sheetname(i) = ela
Next i
sheetname(num_th + 7) = "figure_dyna"
range_Y(0) = "R3C9:R" & NN + 2 & "C9"
name_XY(1) = "层数"

'----------------------------------------------------------------------绘制X向剪力
For i = 0 To num_th + 2
       
    range_X(i) = "R3C" & 11 + 6 * i & ":R" & NN + 2 & "C" & 11 + 6 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

name_XY(0) = "剪力(EX)"

Location(0) = 0 * Width
Location(1) = 0 * Hight

'如果存在反应谱数据，添加正负35%，正负20%反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  For i = 1 To 4
      range_X(num_th + 2 + i) = "R3C" & 11 + 6 * (num_th + 2) + 4 + i & ":R" & NN + 2 & "C" & 11 + 6 * (num_th + 2) + 4 + i
      name_S(num_th + 2 + i) = Sheets(ela).Cells(6 + num_th + 2 + i, 1)
  Next i
  
  Call add_chart_array(sheetname(), num_th + 7, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'----------------------------------------------------------------------绘制Y向剪力
For i = 0 To num_th + 2
    
    range_X(i) = "R3C" & 14 + 6 * i & ":R" & NN + 2 & "C" & 14 + 6 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i
                
name_XY(0) = "剪力(EY)"

Location(0) = 1 * Width
Location(1) = 0 * Hight

'如果存在反应谱数据，添加正负35%，正负20%反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  For i = 1 To 4
      range_X(num_th + 2 + i) = "R3C" & 11 + 6 * (num_th + 2) + 8 + i & ":R" & NN + 2 & "C" & 11 + 6 * (num_th + 2) + 8 + i
      name_S(num_th + 2 + i) = Sheets(ela).Cells(6 + num_th + 2 + i, 1)
  Next i
  
  Call add_chart_array(sheetname(), num_th + 7, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If

'---------------------------------------------------------------------绘制X向弯矩
For i = 0 To num_th + 2
    
    range_X(i) = "R3C" & 12 + 6 * i & ":R" & NN + 2 & "C" & 12 + 6 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

name_XY(0) = "弯矩(EX)"

Location(0) = 0 * Width
Location(1) = 1 * Hight

Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'--------------------------------------------------------------------绘制Y向弯矩
For i = 0 To num_th + 2
    
    range_X(i) = "R3C" & 15 + 6 * i & ":R" & NN + 2 & "C" & 15 + 6 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

name_XY(0) = "弯矩(EY)"

Location(0) = 1 * Width
Location(1) = 1 * Hight

Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'--------------------------------------------------------------------绘制X向层间位移角
For i = 0 To num_th + 2
    
    range_X(i) = "R3C" & 101 + 2 * (num_th + 3) + 2 * i & ":R" & NN + 2 & "C" & 101 + 2 * (num_th + 3) + 2 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i
                
name_XY(0) = "层间位移角(EX)"

Location(0) = 0 * Width
Location(1) = 2 * Hight

Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'----------------------------------------------------------------------绘制Y向层间位移角
For i = 0 To num_th + 2
    
    range_X(i) = "R3C" & 102 + 2 * (num_th + 3) + 2 * i & ":R" & NN + 2 & "C" & 102 + 2 * (num_th + 3) + 2 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

name_XY(0) = "层间位移角(EY)"

Location(0) = 1 * Width
Location(1) = 2 * Hight

Call add_chart_array(sheetname(), num_th + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")


End Sub




