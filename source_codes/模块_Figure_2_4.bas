Attribute VB_Name = "模块_Figure_2_4"

Option Explicit


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                         配筋对比绘图                                 ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/16

'更新内容：
'1.增加比值1线
'2.改为散点图

'////////////////////////////////////////////////////////////////////////////////////////////



'配筋对比数据绘图
'变量说明：
'range_X1,range_X2,range_Y, 为图表的X、Y轴数据range；
'range_X1,range_X2,range_Y,所在的sheet；
'name_X, name_Y为图表的X、Y轴的标题；
'name为数据Series的标题；
' Location_X, Location_Y为图标在sheet中的位置；
'Optional NumFormat As String = "G/通用格式" 为X轴数据的格式，缺省值为通用，可为分数及科学计数等；

Sub add_chart_2(sheetname, range_X1, range_X2, range_Y, name_1, name_2, name_X, name_Y, Location_X, Location_Y, Width, Hight, Optional NumFormat As String = "G/通用格式")

    Debug.Print "开始绘图"
    Debug.Print "..."

    Dim myChart As ChartObject
  
    Dim i, j As Integer
  
    Dim range(), name()
    range = Array(range_X1, range_X2)
    'Debug.Print range(1)
    name = Array(name_1, name_2)
    With Sheets("figure_Info")
                 
        '清除原有图表
        '.ChartObjects.Delete
          
        '指定图表位置和大小.add(左边距，定边距，宽度，高度），该数值单位不是公制,为磅
        Set myChart = .ChartObjects.Add(Location_X, Location_Y, Width, Hight)
      
        '显示边框
        myChart.Border.LineStyle = 1
              
              
'=====================================================================================开始绘图
        With myChart.Chart
      
            '设置绘图区的大小
            With .PlotArea
            On Error Resume Next '此处老出错，忽略不影响
            .Width = Width * 0.9
            .Height = Hight * 0.9
            .Left = Width * 0.08
            .Top = Hight * 0.02
            End With


            '设置图表类型
            '为带平滑线的散点图,如果需要画多种不同类型的曲线，如折线图等，可使用一个Select Case...End Select命令
            .ChartType = xlXYScatterSmoothNoMarkers
          
           '------------------------------------------------------------------------对2个系列循环绘制
           For i = 0 To 1
          
            '添加数据系列
            .SeriesCollection.NewSeries
          
            '选择X轴系列值
'            Debug.Print range(i)
            .SeriesCollection(i + 1).XValues = Sheets(sheetname).range(range(i))

          
            '选择X轴标题字体格式
             With .Axes(xlCategory).TickLabels.Font
                 .name = "Times New Roman"    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
             .Axes(xlCategory).TickLabels.NumberFormatLocal = NumFormat
           
            '选择Y轴系列值
            .SeriesCollection(i + 1).Values = Sheets(sheetname).range(range_Y)
          
          
            '选择Y轴标题字体格式
             With .Axes(xlValue).TickLabels.Font
                 .name = "Times New Roman"    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
                     
          
            '选择系列标题
            .SeriesCollection(i + 1).name = name(i)
            '设置标题字体格式
            .Legend.LegendEntries(i).Font.name = "Times New Roman"
          
            '不显示数据标签值
            .ApplyDataLabels ShowValue:=False
            
                             
            '选择Series的线宽等格式信息
            With .SeriesCollection(i + 1)
                .ChartType = xlXYScatter
                '选择线宽
'                .Format.line.Weight = 2
          
                '选择线颜色 此处不能设定，否则出来的线颜色都一样
                '.Format.line.ForeColor.RGB = RGB(112, 48, 160)
          
                '选择线型-单线
'                .Format.line.Style = msoLineSingle
          
                '选择数据点大小
'                .MarkerStyle = 0
                .MarkerSize = 2
                
'                .Format.line.Visible = msoTrue
          
                '选择线型-短线类型
'                .Format.line.DashStyle = msoLineSolid   '实线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineSysDot  '圆点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineSysDash  '方点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineDash     '短划线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineDashDot  '划线-点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDash  '长划线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDashDot  '长划线-点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDashDotDot '长划线-点-点
          
            End With
            
            
            '显示纵向网格
            .SetElement (msoElementPrimaryCategoryGridLinesMajor)

            '设置纵向网格线宽线型
            With .Axes(xlCategory).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.5
                .DashStyle = msoLineSysDash
            End With
          
            '设置横向网格线宽线型
            With .Axes(xlValue).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.5
                .DashStyle = msoLineSysDash
            End With
              
            '显示X、Y轴刻度
            .HasAxis(xlCategory, xlPrimary) = True  'X轴
            .HasAxis(xlValue, xlPrimary) = True     'Y轴
          
            '设置X、Y轴标题
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            With .Axes(xlValue).AxisTitle
                .Text = name_Y
                .Font.name = "Times New Roman"
                .Font.Size = 10
                .Font.Bold = False
                '.Characters(10, 8).Font.Italic = True
            End With
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            With .Axes(xlCategory).AxisTitle
                .Text = name_X
                .Font.name = "Times New Roman"
                .Font.Size = 10
                .Font.Bold = False
                '.Characters(10, 8).Font.Italic = True
            End With
             
            '不显示显示图表标题，若写，可在sheet的对应位置写，不必占用图表空间
            .HasTitle = False
            '.ChartTitle.Text = "刚度比"
            'With .ChartTitle.Font
                '.Size = 20
                '.ColorIndex = 1
                '.Name = "华文新魏"
            'End With
                      
            '.设置图表区域的颜色，缺省就用白色
            'With .ChartArea.Interior
                '.ColorIndex = 2
                '.PatternColorIndex = 1
                '.Pattern = xlSolid
            'End With
          
            '.设置绘图区域的颜色
            With .PlotArea.Interior
                .ColorIndex = 20
                .PatternColorIndex = 1
                .Pattern = xlSolid
            End With
            
            
        Next
        
'==============================================================================================================绘制比值1
        '添加数据系列
        .SeriesCollection.NewSeries
                                       
        '选择X轴系列值
        .SeriesCollection(3).XValues = "={1,1}"
                
        '选择Y轴系列值
        Dim n_y As Integer
        n_y = Sheets(sheetname).range(range_Y).Cells.Count
        Debug.Print n_y
        Debug.Print Sheets(sheetname).range(range_Y)
        .SeriesCollection(3).Values = "={1," & n_y & "}"
                                   
        '选择系列标题
        .SeriesCollection(3).name = "比值1"
                
        With .SeriesCollection(3)
            '选择线宽
            .Format.Line.Weight = 2
                
            '选择线颜色
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                
            '选择线型-单线
            .Format.Line.Style = msoLineSingle
                
            '选择线型-短线类型
            .Format.Line.DashStyle = msoLineSolid   '实线
                    
            '选择数据点类型
            .MarkerStyle = 0
                    
        End With
        
        
        .Legend.Select
        Selection.Font.name = "Times New Roman"

        End With
      
        '.清空对象
        Set myChart = Nothing
          
    End With

    Debug.Print "绘图结束"

End Sub


Sub add_chart_4(sheetname, range_X1, range_X2, range_X3, range_X4, range_Y, name_1, name_2, name_3, name_4, name_X, name_Y, Location_X, Location_Y, Width, Hight, Optional NumFormat As String = "G/通用格式")

    Debug.Print "开始绘图"
    Debug.Print "..."

    Dim myChart As ChartObject
  
    Dim i, j As Integer
  
    Dim range(), name()
    range = Array(range_X1, range_X2, range_X3, range_X4)
    'Debug.Print range(1)
    name = Array(name_1, name_2, name_3, name_4)
    With Sheets("figure_Info")
                 
        '清除原有图表
        '.ChartObjects.Delete
          
        '指定图表位置和大小.add(左边距，定边距，宽度，高度），该数值单位不是公制,为磅
        Set myChart = .ChartObjects.Add(Location_X, Location_Y, Width, Hight)
      
        '显示边框
        myChart.Border.LineStyle = 1
              
              
'=====================================================================================开始绘图
        With myChart.Chart
      
            '设置绘图区的大小
            With .PlotArea
            On Error Resume Next '此处老出错，忽略不影响
            .Width = Width * 0.9
            .Height = Hight * 0.9
            .Left = Width * 0.08
            .Top = Hight * 0.02
            End With


            '设置图表类型
            '为带平滑线的散点图,如果需要画多种不同类型的曲线，如折线图等，可使用一个Select Case...End Select命令
            .ChartType = xlXYScatterSmoothNoMarkers
          
           '------------------------------------------------------------------------对4个系列循环绘制
           For i = 0 To 3
          
            '添加数据系列
            .SeriesCollection.NewSeries
          
            '选择X轴系列值
'            Debug.Print range(i)
            .SeriesCollection(i + 1).XValues = Sheets(sheetname).range(range(i))

          
            '选择X轴标题字体格式
             With .Axes(xlCategory).TickLabels.Font
                 .name = "Times New Roman"    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
             .Axes(xlCategory).TickLabels.NumberFormatLocal = NumFormat
           
            '选择Y轴系列值
            .SeriesCollection(i + 1).Values = Sheets(sheetname).range(range_Y)
          
          
            '选择Y轴标题字体格式
             With .Axes(xlValue).TickLabels.Font
                 .name = "Times New Roman"    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
                     
          
            '选择系列标题
            .SeriesCollection(i + 1).name = name(i)
            '设置标题字体格式
            .Legend.LegendEntries(i).Font.name = "Times New Roman"
          
            '不显示数据标签值
            .ApplyDataLabels ShowValue:=False
            
                             
            '选择Series的线宽等格式信息
            With .SeriesCollection(i + 1)
                .ChartType = xlXYScatter
                '选择线宽
'                .Format.line.Weight = 1
'
'                '选择线颜色 此处不能设定，否则出来的线颜色都一样
'                '.Format.line.ForeColor.RGB = RGB(112, 48, 160)
'
'                '选择线型-单线
'                .Format.line.Style = msoLineSingle
'
'                '选择数据点大小
'                .MarkerStyle = 0
                .MarkerSize = 2
'
'                .Format.line.Visible = msoTrue
'
'                '选择线型-短线类型
'                .Format.line.DashStyle = msoLineSolid   '实线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineSysDot  '圆点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineSysDash  '方点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineDash     '短划线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineDashDot  '划线-点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDash  '长划线
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDashDot  '长划线-点
            '.SeriesCollection(1).Format.Line.DashStyle = msoLineLongDashDotDot '长划线-点-点
          
            End With
            
            
            '显示纵向网格
            .SetElement (msoElementPrimaryCategoryGridLinesMajor)

            '设置纵向网格线宽线型
            With .Axes(xlCategory).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.5
                .DashStyle = msoLineSysDash
            End With
          
            '设置横向网格线宽线型
            With .Axes(xlValue).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.5
                .DashStyle = msoLineSysDash
            End With
              
            '显示X、Y轴刻度
            .HasAxis(xlCategory, xlPrimary) = True  'X轴
            .HasAxis(xlValue, xlPrimary) = True     'Y轴
          
            '设置X、Y轴标题
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            With .Axes(xlValue).AxisTitle
                .Text = name_Y
                .Font.name = "Times New Roman"
                .Font.Size = 10
                .Font.Bold = False
                '.Characters(10, 8).Font.Italic = True
            End With
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            With .Axes(xlCategory).AxisTitle
                .Text = name_X
                .Font.name = "Times New Roman"
                .Font.Size = 10
                .Font.Bold = False
                '.Characters(10, 8).Font.Italic = True
            End With
             
            '不显示显示图表标题，若写，可在sheet的对应位置写，不必占用图表空间
            .HasTitle = False
            '.ChartTitle.Text = "刚度比"
            'With .ChartTitle.Font
                '.Size = 20
                '.ColorIndex = 1
                '.Name = "华文新魏"
            'End With
                      
            '.设置图表区域的颜色，缺省就用白色
            'With .ChartArea.Interior
                '.ColorIndex = 2
                '.PatternColorIndex = 1
                '.Pattern = xlSolid
            'End With
          
            '.设置绘图区域的颜色
            With .PlotArea.Interior
                .ColorIndex = 20
                .PatternColorIndex = 1
                .Pattern = xlSolid
            End With
            
            
        Next
        
'==============================================================================================================绘制比值1
        '添加数据系列
        .SeriesCollection.NewSeries
                                       
        '选择X轴系列值
        .SeriesCollection(5).XValues = "={1,1}"
                
        '选择Y轴系列值
        Dim n_y As Integer
        n_y = Sheets(sheetname).range(range_Y).Cells.Count
        Debug.Print n_y
        Debug.Print Sheets(sheetname).range(range_Y)
        .SeriesCollection(5).Values = "={1," & n_y & "}"
                                   
        '选择系列标题
        .SeriesCollection(5).name = "比值1"
                
        With .SeriesCollection(5)
            '选择线宽
            .Format.Line.Weight = 2
                
            '选择线颜色
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                
            '选择线型-单线
            .Format.Line.Style = msoLineSingle
                
            '选择线型-短线类型
            .Format.Line.DashStyle = msoLineSolid   '实线
                    
            '选择数据点类型
            .MarkerStyle = 0
                    
        End With
        
        
        
        .Legend.Select
        Selection.Font.name = "Times New Roman"

        End With
      
        '.清空对象
        Set myChart = Nothing
          
    End With

    Debug.Print "绘图结束"

End Sub

