Attribute VB_Name = "ETABS_HIST_Figs"
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                       ETABS时程数据画图                              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改
'2.修正Y向时程数少于X向时时程曲线出错bug

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/8
'更新内容:
'1.由于X向和Y向时程分开了，所以不能调用原先的时程数据画图代码，只能单独写了，可以想象办法统一一下

Sub ETABS_HIST_Fig(ela As String)

'Dim ela  As String
'ela = "e_E"
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
Dim num_th, num_th1, Num_X, Num_Y As Integer
num_th = Sheets(ela).Cells(2, 2)
num_th1 = num_th
Num_X = Sheets(ela).Cells(2, 4)
Num_Y = Sheets(ela).Cells(2, 6)
'Debug.Print num_th

'如有反应谱数据，绘制曲线数减2
If Sheets(ela).Cells(num_th + 8, 2) = "" Then
    num_th = num_th - 2
End If

For i = 0 To num_th + 6 - 1
    Sheets(ela).range(Sheets(ela).Cells(3, 100 + i), Sheets(ela).Cells(NN + 2, 100 + i)).Value = _
    Sheets(ela).range(Sheets(ela).Cells(3, 10 + 3 * i), Sheets(ela).Cells(NN + 2, 10 + 3 * i)).Value
Next

With Sheets(ela).range(Sheets(ela).Cells(3, 100 + num_th + 6 + 1), Sheets(ela).Cells(NN + 2, 100 + 2 * (num_th + 6)))
    .FormulaR1C1 = "=1/RC[-" & num_th + 6 + 1 & "]"
    .Font.ColorIndex = 2
    .Locked = True
End With

With Sheets(ela).range(Sheets(ela).Cells(3, 100), Sheets(ela).Cells(NN + 2, 100 + 2 * (num_th + 3)))
    .Font.ColorIndex = 2
    .Locked = True
End With


Sheets("figure_dyna").Select

'--------------------------------------------------------------绘图公用部分
Dim sheetname(), range_X(), range_Y(0), name_S(), name_XY(1)
ReDim Preserve sheetname(num_th + 7)
ReDim Preserve range_X(num_th1 + 6)
ReDim Preserve name_S(num_th1 + 6)
Dim Location(1) As Integer

'定义图表高宽
Dim Width As Integer, Hight As Integer

Width = 414
Hight = 510

For i = 0 To num_th + 6
    sheetname(i) = ela
Next i
sheetname(num_th + 7) = "figure_dyna"
range_Y(0) = "R3C9:R" & NN + 2 & "C9"
name_XY(1) = "层数"


'--------------------------------------------------------------绘制X向剪力
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_X - 1
       
    range_X(i) = "R3C" & 11 + 3 * i & ":R" & NN + 2 & "C" & 11 + 3 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_X + i) = "R3C" & 11 + 3 * (num_th1 + 2 * i) & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 2 * i)
         
    name_S(Num_X + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "剪力(EX)"

Location(0) = 0 * Width
Location(1) = 0 * Hight

'如果存在反应谱数据，添加正负35%，正负20%反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), Num_X + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_X + 2) = "R3C" & 11 + 3 * (num_th1 + 4) & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 4)
  name_S(Num_X + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)
      
  For i = 1 To 4
      range_X(Num_X + 2 + i) = "R3C" & 11 + 3 * (num_th1 + 6) - 2 + i & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 6) - 2 + i
      name_S(Num_X + 2 + i) = Sheets(ela).Cells(6 + num_th1 + 2 + i, 1)
  Next i
  
  Call add_chart_array(sheetname(), Num_X + 7, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'-------------------------------------------------------------绘制Y向剪力
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_Y - 1
       
    range_X(i) = "R3C" & 11 + 3 * (Num_X + i) & ":R" & NN + 2 & "C" & 11 + 3 * (Num_X + i)
         
    name_S(i) = Sheets(ela).Cells(6 + Num_X + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_Y + i) = "R3C" & 11 + 3 * (num_th1 + 1 + 2 * i) & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 1 + 2 * i)
         
    name_S(Num_Y + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "剪力(EY)"

Location(0) = 1 * Width
Location(1) = 0 * Hight

'如果存在反应谱数据，添加正负35%，正负20%反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 5) = "" Then

  Call add_chart_array(sheetname(), Num_Y + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_Y + 2) = "R3C" & 11 + 3 * (num_th1 + 5) & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 5)
  name_S(Num_Y + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)
      
  For i = 1 To 4
      range_X(Num_Y + 2 + i) = "R3C" & 11 + 3 * (num_th1 + 6) + 2 + i & ":R" & NN + 2 & "C" & 11 + 3 * (num_th1 + 6) + 2 + i
      name_S(Num_Y + 2 + i) = Sheets(ela).Cells(6 + num_th1 + 2 + i, 1)
  Next i
  
  Call add_chart_array(sheetname(), Num_Y + 7, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'-------------------------------------------------------------绘制X向弯矩
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_X - 1
       
    range_X(i) = "R3C" & 12 + 3 * i & ":R" & NN + 2 & "C" & 12 + 3 * i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_X + i) = "R3C" & 12 + 3 * (num_th1 + 2 * i) & ":R" & NN + 2 & "C" & 12 + 3 * (num_th1 + 2 * i)
         
    name_S(Num_X + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "弯矩(EX)"

Location(0) = 0 * Width
Location(1) = 1 * Hight

'如果存在反应谱数据，添加反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), Num_X + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_X + 2) = "R3C" & 12 + 3 * (num_th1 + 4) & ":R" & NN + 2 & "C" & 12 + 3 * (num_th1 + 4)
  name_S(Num_X + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)

  Call add_chart_array(sheetname(), Num_X + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'-------------------------------------------------------------绘制Y向弯矩
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_Y - 1
       
    range_X(i) = "R3C" & 12 + 3 * (Num_X + i) & ":R" & NN + 2 & "C" & 12 + 3 * (Num_X + i)
         
    name_S(i) = Sheets(ela).Cells(6 + Num_X + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_Y + i) = "R3C" & 12 + 3 * (num_th1 + 1 + 2 * i) & ":R" & NN + 2 & "C" & 12 + 3 * (num_th1 + 1 + 2 * i)
         
    name_S(Num_Y + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "弯矩(EY)"

Location(0) = 1 * Width
Location(1) = 1 * Hight

'如果存在反应谱数据，添加反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), Num_Y + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_Y + 2) = "R3C" & 12 + 3 * (num_th1 + 5) & ":R" & NN + 2 & "C" & 12 + 3 * (num_th1 + 5)
  name_S(Num_Y + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)

  Call add_chart_array(sheetname(), Num_Y + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'-----------------------------------------------------------绘制X向层间位移角
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_X - 1
       
    range_X(i) = "R3C" & 107 + num_th + i & ":R" & NN + 2 & "C" & 107 + num_th + i
         
    name_S(i) = Sheets(ela).Cells(6 + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_X + i) = "R3C" & 107 + num_th + num_th1 + 2 * i & ":R" & NN + 2 & "C" & 107 + num_th + num_th1 + 2 * i
         
    name_S(Num_X + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "层间位移角(EX)"

Location(0) = 0 * Width
Location(1) = 2 * Hight

'如果存在反应谱数据，添加反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), Num_X + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_X + 2) = "R3C" & 107 + num_th + num_th1 + 4 & ":R" & NN + 2 & "C" & 107 + num_th + num_th1 + 4
  name_S(Num_X + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)

  Call add_chart_array(sheetname(), Num_X + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


'-----------------------------------------------------------绘制Y向层间位移角
Erase range_X
ReDim Preserve range_X(num_th1 + 6)

For i = 0 To Num_Y - 1
       
    range_X(i) = "R3C" & 107 + num_th + Num_X + i & ":R" & NN + 2 & "C" & 107 + num_th + Num_X + i
         
    name_S(i) = Sheets(ela).Cells(6 + Num_X + i, 1)
    
Next i

'平均值和最大值
For i = 0 To 1
       
    range_X(Num_Y + i) = "R3C" & 108 + num_th + num_th1 + 2 * i & ":R" & NN + 2 & "C" & 108 + num_th + num_th1 + 2 * i
         
    name_S(Num_Y + i) = Sheets(ela).Cells(6 + num_th1 + i, 1)
    
Next i

name_XY(0) = "层间位移角(EY)"

Location(0) = 1 * Width
Location(1) = 2 * Hight

'如果存在反应谱数据，添加反应谱剪力曲线
If Sheets(ela).Cells(num_th1 + 8, 2) = "" Then

  Call add_chart_array(sheetname(), Num_Y + 2, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

Else

  range_X(Num_Y + 2) = "R3C" & 108 + num_th + num_th1 + 4 & ":R" & NN + 2 & "C" & 108 + num_th + num_th1 + 4
  name_S(Num_Y + 2) = Sheets(ela).Cells(6 + num_th1 + 2, 1)

  Call add_chart_array(sheetname(), Num_Y + 3, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)
  
End If


End Sub


