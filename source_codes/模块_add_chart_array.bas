Attribute VB_Name = "模块_add_chart_array"
Option Explicit

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                         绘制多列数据图                               ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/4/20
'1.将Y轴分隔改为10层


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.针对多模型对比修改限值代码
'2.线型改变为实线
'3.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/8
'更新内容：
'1.无法却别小震和大震的时程分析层间位移角限值，取消限值画图（之前一直用它画大震时程图，所以添加了，现在看没用了）

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/7
'更新内容：
'1.增加时程分析层间位移角限值

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/5

'更新内容：
'1.将“剪重比”分成X、Y两个；


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27

'更新内容：
'1.改为使用数组传递进行绘图，通用性较好
'2.更改模块名与过程名相同


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/25

'更新内容：
'1.隐去.Format.line.ForeColor.RGB设定，使颜色分开；
'2..MarkerStyle = I +改为.MarkerStyle = I + 1，使得第一个数据为点线格式；


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/18 22:28

'更新内容：
'1.添加点线线型语句


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/24

'更新内容：
'1.添加说明；

'时程分析数据绘图
'变量说明：
'sheetname() 为曲线数据所在sheet名，1 X N_C+1 维数组，最后一个值为图表所在sheet名
'N_C 为图中曲线总数
'range_X() 为图表的X轴数据range，1 X N_C 维数组
'range_Y() 为图表的Y轴数据range，1 X 1 维数组
'name() 为数据Series的标题，1 X N_C 维数组
'name_XY() 1 X 2 维数组，第一个值为图表的X轴的标题，第二个值为Y轴标题
'Location() 1 X 2 维数组，为图标在sheet中的位置；
'Optional NumFormat As String = "G/通用格式" 为X轴数据的格式，缺省值为通用，可为分数及科学计数等；

Sub add_chart_array(sheetname(), N_C, range_X(), range_Y(), name(), name_XY(), Location() As Integer, Width, Hight, Optional NumFormat As String = "G/通用格式")

    Debug.Print "开始绘图"
    Debug.Print "..."
   
    '统一设定字体
    Dim ft As String
    ft = "Arial"

    Dim myChart As ChartObject
     
    Dim i, j As Integer

    With Sheets(sheetname(UBound(sheetname())))
                
        '清除原有图表
        '.ChartObjects.Delete
         
        '指定图表位置和大小.add(左边距，定边距，宽度，高度），该数值单位不是公制,为磅
        Set myChart = .ChartObjects.Add(Location(0), Location(1), Width, Hight)
     
        '显示边框
        myChart.Border.LineStyle = 1
             
             
'=====================================================================================开始绘图
        With myChart.Chart
       
            .ChartArea.Format.Line.Visible = msoFalse '---------------去掉外围线
     
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
         
           '------------------------------------------------------------------------对10个系列循环绘制
           For i = 0 To N_C - 1
         
            '添加数据系列
            .SeriesCollection.NewSeries
         
            '选择X轴系列值

            .SeriesCollection(i + 1).XValues = sheetname(i) & "!" & range_X(i)
            '.SeriesCollection(i + 1).XValues = Sheets(sheetname).range(range_X(i))
         
            '选择X轴标题字体格式
             With .Axes(xlCategory).TickLabels.Font
                 .name = ft    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
             .Axes(xlCategory).TickLabels.NumberFormatLocal = NumFormat
            
             '设置X轴刻度线
             .Axes(xlCategory).MajorTickMark = xlInside
              With .Axes(xlCategory).Format.Line
                .ForeColor.RGB = RGB(153, 76, 0)
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Weight = 0.75
              End With
          
            '选择Y轴系列值
            .SeriesCollection(i + 1).Values = sheetname(i) & "!" & range_Y(0)
            '.SeriesCollection(i + 1).Values = Sheets(sheetname).range(range_Y(0))
         
            '选择Y轴标题字体格式
             With .Axes(xlValue).TickLabels.Font
                 .name = ft    '双引号中间填写你需要的字体
                 '.FontStyle = ""   '是否加粗等格式
                 .Size = 9   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
            
             '设置Y轴刻度线
             .Axes(xlValue).MajorTickMark = xlInside 'xlnone
              With .Axes(xlValue).Format.Line
                .ForeColor.RGB = RGB(153, 76, 0)
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Weight = 0.75
              End With
             
              '设定Y轴坐标轴上限(绘图时间会增加6%)
              .Axes(xlValue).MaximumScale = Int(Num_all / 10 + 1) * 10
                    
         
            '选择系列标题
            .SeriesCollection(i + 1).name = name(i)
            '设置标题字体格式
            .Legend.LegendEntries(i).Font.name = ft
            .Legend.LegendEntries(i).Font.Size = 9
            .Legend.Format.TextFrame2.TextRange.Font.Size = 9
            .Legend.Format.Fill.Visible = msoTrue
            .Legend.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            With .Legend.Format.Line
                .Visible = msoTrue
                 .ForeColor.RGB = RGB(0, 0, 0)
            End With
           
            '.Legend.Position = xlLegendPositionTop
            '.Legend.Top = 0
'            .Legend.Format.
'    With Selection.Format.Line
'        .Visible = msoTrue
'        .ForeColor.RGB = RGB(238, 236, 225)
'    End With
'    ActiveChart.Axes(xlCategory).Select
'    With Selection.Format.Fill
'        .Visible = msoTrue
'        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
'        .ForeColor.TintAndShade = 0
'        .ForeColor.Brightness = 0
'        .Transparency = 0
'        .Solid
'    End With
           
            '不显示数据标签值
            .ApplyDataLabels ShowValue:=False
           
                            
            '选择Series的线宽等格式信息
            With .SeriesCollection(i + 1)
                '选择线宽
                .Format.Line.Weight = 1.5
                If i = 0 Then
                 .Format.Line.ForeColor.RGB = RGB(0, 153, 0) '(204, 51, 0)
                ElseIf i = 1 Then
                 .Format.Line.ForeColor.RGB = RGB(0, 102, 255)
                End If
         
                '选择线颜色 此处不能设定，否则出来的线颜色都一样
                '.Format.line.ForeColor.RGB = RGB(112, 48, 160)
         
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
         
                '选择数据点大小
                .MarkerStyle = 0
                '.MarkerStyle = i + 1
                '.MarkerSize = 2.5
               
                .Format.Line.Visible = msoTrue
         
                '选择线型-短线类型
                .Format.Line.DashStyle = msoLineSolid   '实线
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
                .Weight = 0.1
                .DashStyle = msoLineDash
                .ForeColor.RGB = RGB(0, 0, 0)
            End With
         
            '设置横向网格线宽线型
            With .Axes(xlValue).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.1
                .DashStyle = msoLineDash
                .ForeColor.RGB = RGB(0, 0, 0)
            End With
             
            '显示X、Y轴刻度
            .HasAxis(xlCategory, xlPrimary) = True  'X轴
            .HasAxis(xlValue, xlPrimary) = True     'Y轴
         
            '设置X、Y轴标题
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            With .Axes(xlValue).AxisTitle
                .Text = name_XY(1)
                .Font.name = ft
                .Font.Size = 9
                .Font.Bold = True
                '.Characters(10, 8).Font.Italic = True
            End With
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            With .Axes(xlCategory).AxisTitle
                .Text = name_XY(0)
                .Font.name = ft
                .Font.Size = 9
                .Font.Bold = True
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
         
            '.设置绘图区域的颜色，缺省就用白色
'            With .PlotArea.Interior
'                .ColorIndex = 20
'                .PatternColorIndex = 1
'                .Pattern = xlSolid
'            End With
'.Line
            ' .设置绘图区域的颜色，缺省就用白色
            With .PlotArea.Format.Line
                .Visible = msoTrue
                .Weight = 1
                .DashStyle = msoLineSolid
                .ForeColor.RGB = RGB(0, 0, 0)
            End With
                       
        Next
       
        '-------------------------------------------------------------时程分析位移角限值
'        Dim Temp_String As String
'       Temp_String = name_XY(0)
       
  '      If CheckRegExpfromString(Temp_String, "层间位移角") Then
                            
                '添加数据系列
   '             .SeriesCollection.NewSeries
                                       
                '选择X轴系列值、
    '            .SeriesCollection(N_C + 1).XValues = "={ 0.01，0.01}"
                
                 '选择Y轴系列值
     '           .SeriesCollection(N_C + 1).Values = "={0," & Num_All & "}"
                                   
                '选择系列标题
      '          .SeriesCollection(N_C + 1).name = "规范限值1/100"
           
       '         With .SeriesCollection(N_C + 1)
                    '选择线宽
        '            .Format.Line.Weight = 2
               
                    '选择线颜色
         '           .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
               
                    '选择线型-单线
          '          .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
           '         .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
            '        .MarkerStyle = 0
                   
             '   End With
       
        'End If
       
        '-------------------------------------------------------------位移比限值
        Dim name11 As String
        name11 = name_XY(0)
        If CheckRegExpfromString(name11, "位移比") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={1.4,1.4}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值1.4"
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If
       
        '-------------------------------------------------------------位移角限值
        If CheckRegExpfromString(name11, "下位移角") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={" & 1 / OUTReader_Main.DisLimit_TextBox.Text & "," & 1 / OUTReader_Main.DisLimit_TextBox.Text & "}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值1/" & OUTReader_Main.DisLimit_TextBox.Text
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If

        '-------------------------------------------------------------X剪重比限值
        If CheckRegExpfromString(name11, "X向剪重比") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={" & OUTReader_Main.RatioLimitX_TextBox.Text & "," & OUTReader_Main.RatioLimitX_TextBox.Text & "}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值(" & OUTReader_Main.RatioLimitX_TextBox.Text
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If
       
        '-------------------------------------------------------------Y剪重比限值
        If CheckRegExpfromString(name11, "Y向剪重比") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={" & OUTReader_Main.RatioLimitY_TextBox.Text & "," & OUTReader_Main.RatioLimitY_TextBox.Text & "}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值(" & OUTReader_Main.RatioLimitY_TextBox.Text
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If
        '-------------------------------------------------------------承载力比限值
        If CheckRegExpfromString(name11, "承载力") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={0.75,0.75}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值0.75"
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If
       
        '-------------------------------------------------------------质量比限值
        If CheckRegExpfromString(name11, "质量比") Then
       
           
                '添加数据系列
                .SeriesCollection.NewSeries
                                      
                '选择X轴系列值
                .SeriesCollection(N_C + 1).XValues = "={1.5,1.5}"
               
                '选择Y轴系列值
                .SeriesCollection(N_C + 1).Values = "={0," & Num_all & "}"
                                  
                '选择系列标题
                .SeriesCollection(N_C + 1).name = "限值1.5"
               
                With .SeriesCollection(N_C + 1)
                    '选择线宽
                    .Format.Line.Weight = 1.5
               
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
               
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
               
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                   
                    '选择数据点类型
                    .MarkerStyle = 0
                   
                End With
       
        End If
       
       
        .Legend.Select
        Selection.Font.name = ft

        End With
     
        '.清空对象
        Set myChart = Nothing
         
    End With

    Debug.Print "绘图结束"

End Sub


