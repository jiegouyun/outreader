Attribute VB_Name = "模块_add_chart"
Option Explicit

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                           绘制单列数据图                             ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/25
'1.更新位移比限值画图，在表g_*中填写限值则画图取该值，不填写则规范1.2，1.4都画出；
'2.纵坐标加上限，位移比横坐标从1开始

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27
'1.更改为三种软件通用型，根据输入参数选择绘图；
'2.更改模块名与过程名相同

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/18

'更新内容：
'1.添加位移比,刚度比,剪重比,质量比,抗剪承载力比限值曲线.
'2.层间位移角因结构形式和莪结构高度原因未能确定，两种解决方案：1，手动输入，（目前选择），2，添加判断（麻烦）

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/6

'更新内容：
'1.添加说明；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/5

'更新内容：
'1.打开图标外框线，更美观些；


'///////////////////////////////////////////////////////////
'这只是绘图的模板，具体绘图时需要改进
'我设想正式编写绘图程序时，可以设置函数引用变量，画图的主程序是不变的，根据输入的不同变量画不同的图。
'///////////////////////////////////////////////////////////



'读取楼层数进行专项绘图
'变量说明：
'softname 为软件名称
'range_X, range_Y为图表的X、Y轴数据range；
'dis_sheet为range_X, range_Y所在的sheet；
'name_X, name_Y为图表的X、Y轴的标题；
'name为数据Series的标题；
' Location_X, Location_Y为图标在sheet中的位置；
'Optional NumFormat As String = "G/通用格式" 为X轴数据的格式，缺省值为通用，可为分数及科学计数等；

Sub add_chart(softname, range_X, range_Y, name, name_X, name_Y, Location_X, Location_Y, Width, Hight, Optional NumFormat As String = "G/通用格式")

    Debug.Print "开始绘图"
    Debug.Print "..."

    Dim myChart As ChartObject
    
    Dim i, j As Integer
    
    Dim dis_sheet, gen_sheet, fig_sheet As String
    
    If softname = "PKPM" Then
        dis_sheet = "d_P"
        gen_sheet = "g_P"
        fig_sheet = "figure_PKPM"
    ElseIf softname = "YJK" Then
        dis_sheet = "d_Y"
        gen_sheet = "g_Y"
        fig_sheet = "figure_YJK"
    ElseIf softname = "MBuilding" Then
        dis_sheet = "d_M"
        gen_sheet = "g_M"
        fig_sheet = "figure_MBuilding"
    ElseIf softname = "ETABS" Then
        dis_sheet = "d_E"
        gen_sheet = "g_E"
        fig_sheet = "figure_ETABS"
    Else
        MsgBox "参数输入错误"
    End If
    
    With Sheets(fig_sheet)
            
        '指定图表位置和大小.add(左边距，定边距，宽度，高度），该数值单位不是公制,为磅
        Set myChart = .ChartObjects.Add(Location_X, Location_Y, Width, Hight)
        
        '显示边框
        myChart.Border.LineStyle = 1
                
                
'============================开始绘图==========================
        
        With myChart.Chart
        
            '设置绘图区的大小
            .PlotArea.Select
            Selection.Width = Width * 0.9
            Selection.Height = Hight * 0.9
            Selection.Left = Width * 0.08
            Selection.Top = Hight * 0.02


            '设置图表类型为带平滑线的散点图,如果需要画多种不同类型的曲线，如折线图等，可使用一个Select Case...End Select命令
            .ChartType = xlXYScatterSmoothNoMarkers
            
            '添加数据系列
            .SeriesCollection.NewSeries
            
            '选择X轴系列值
            .SeriesCollection(1).XValues = Sheets(dis_sheet).range(range_X)
            
            '选择X轴标题字体格式
             With .Axes(xlCategory).TickLabels.Font
                 .name = "Arial"    '双引号中间填写你需要的字体
                 '.FontStyle = "Bold"   '是否加粗等格式
                 .Size = 10   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
             .Axes(xlCategory).TickLabels.NumberFormatLocal = NumFormat
             
             '设置X轴刻度线
             .Axes(xlCategory).MajorTickMark = xlNone
              With .Axes(xlCategory).Format.Line
                .ForeColor.RGB = RGB(0, 112, 192)
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Weight = 1
              End With
             
            '选择Y轴系列值
            .SeriesCollection(1).Values = Sheets(dis_sheet).range(range_Y)
                        
            '选择Y轴标题字体格式
             With .Axes(xlValue).TickLabels.Font
                 .name = "Arial"    '双引号中间填写你需要的字体
                 '.FontStyle = "Bold"   '是否加粗等格式
                 .Size = 10   '选择字体大小
                 .ColorIndex = 1    '字体颜色
             End With
             
             '设置Y轴刻度线
             .Axes(xlValue).MajorTickMark = xlNone
              With .Axes(xlValue).Format.Line
                .ForeColor.RGB = RGB(0, 112, 192)
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .Weight = 1
              End With
                       
              '设定Y轴坐标轴上限(绘图时间会增加6%)
              .Axes(xlValue).MaximumScale = Int(Num_all / 5 + 1) * 5
            
            '选择系列标题
            .SeriesCollection(1).name = name
            .Legend.Select
            Selection.Font.name = "Arial"
            
            '不显示数据标签值
            .ApplyDataLabels ShowValue:=False
                               
            '选择Series的线宽等格式信息
            With .SeriesCollection(1)
                '选择线宽
                .Format.Line.Weight = 2
            
                '选择线颜色
                .Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
            
                '选择数据点大小
                .MarkerStyle = 0
                '.MarkerSize = 2
            
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
                .Weight = 0.25
                .DashStyle = msoLineDash
            End With
            
            '设置横向网格线宽线型
            With .Axes(xlValue).MajorGridlines.Format.Line
                .Visible = msoTrue
                .Weight = 0.25
                .DashStyle = msoLineDash
            End With
                
            '显示X、Y轴刻度
            .HasAxis(xlCategory, xlPrimary) = True  'X轴
            .HasAxis(xlValue, xlPrimary) = True     'Y轴
            
            '设置X、Y轴标题
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            With .Axes(xlValue).AxisTitle
                .Text = name_Y
                .Font.name = "Arial"
                .Font.Size = 10
                .Font.Bold = True
                '.Characters(10, 8).Font.Italic = True
            End With
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            With .Axes(xlCategory).AxisTitle
                .Text = name_X
                .Font.name = "Arial"
                .Font.Size = 10
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
            
            '.设置绘图区域的颜色
            With .PlotArea.Interior
                .ColorIndex = 20
                .PatternColorIndex = 1
                .Pattern = xlSolid
            End With

        '-------------------------------------------------------------位移比限值
        If name_X = "位移比" Then
        
            'X轴从1开始
            .Axes(xlCategory).MinimumScale = 1
        
            If Not IsEmpty(Sheets(gen_sheet).Cells(16, 7)) Then
            
                '添加数据系列
                .SeriesCollection.NewSeries
                                       
                '选择X轴系列值
                .SeriesCollection(2).XValues = "={" & Sheets(gen_sheet).Cells(16, 7) & "," & Sheets(gen_sheet).Cells(16, 7) & "}"
                
                '选择Y轴系列值
                .SeriesCollection(2).Values = "={0," & Num_all & "}"
                                   
                '选择系列标题
                .SeriesCollection(2).name = "限值" & Sheets(gen_sheet).Cells(16, 7)
                
                With .SeriesCollection(2)
                    '选择线宽
                    .Format.Line.Weight = 2
                
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
                
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
                
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                    
                    '选择数据点类型
                    .MarkerStyle = 0
                    
                End With
            
            Else
            
                '添加数据系列
                .SeriesCollection.NewSeries
                                       
                '选择X轴系列值
                '.SeriesCollection(2).XValues = "={" & Sheets(gen_sheet).Cells(16, 7) & "," & Sheets(gen_sheet).Cells(16, 7) & "}"
                .SeriesCollection(2).XValues = "={1.2,1.2}"
                
                '选择Y轴系列值
                .SeriesCollection(2).Values = "={0," & Num_all & "}"
                                   
                '选择系列标题
                .SeriesCollection(2).name = "限值1.2"
                
                With .SeriesCollection(2)
                    '选择线宽
                    .Format.Line.Weight = 2
                
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
                
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
                
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                    
                    '选择数据点类型
                    .MarkerStyle = 0
                    
                End With
                
                '添加数据系列
                .SeriesCollection.NewSeries
                                       
                '选择X轴系列值
                .SeriesCollection(3).XValues = "={1.4,1.4}"
                
                '选择Y轴系列值
                .SeriesCollection(3).Values = "={0," & Num_all & "}"
                                   
                '选择系列标题
                .SeriesCollection(3).name = "限值1.4"
                
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
            
            End If
        
        End If
        
        
        '-------------------------------------------------------------位移角限值
        If name_X = "位移角" Then
        
            If Not IsEmpty(Sheets(gen_sheet).Cells(14, 7)) Then
            
                '添加数据系列
                .SeriesCollection.NewSeries
                                        
                '选择X轴系列值、
                .SeriesCollection(2).XValues = "={" & Sheets(gen_sheet).Cells(14, 7) & "," & Sheets(gen_sheet).Cells(14, 7) & "}"
                 
                 '选择Y轴系列值
                .SeriesCollection(2).Values = "={0," & Num_all & "}"
                                    
                '选择系列标题
                .SeriesCollection(2).name = "规范限值"
            
                With .SeriesCollection(2)
                    '选择线宽
                    .Format.Line.Weight = 2
                
                    '选择线颜色
                    .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
                
                    '选择线型-单线
                    .Format.Line.Style = msoLineSingle
                
                    '选择线型-短线类型
                    .Format.Line.DashStyle = msoLineSolid   '实线
                    
                    '选择数据点类型
                    .MarkerStyle = 0
                    
                End With
            
            End If
        
        End If
        
        
        '-------------------------------------------------------------剪重比限值
        If name = "X向剪重比" Then
        
            
            If Not IsEmpty(Sheets(gen_sheet).Cells(24, 7)) Then
            
                    '添加数据系列
                    .SeriesCollection.NewSeries
                    
                    '选择X轴系列值
                    .SeriesCollection(2).XValues = "={" & Sheets(gen_sheet).Cells(24, 7) & "," & Sheets(gen_sheet).Cells(24, 7) & "}"
                    
                    '选择Y轴系列值
                    .SeriesCollection(2).Values = "={0," & Num_all & "}"
                    
                    '选择系列标题
                    .SeriesCollection(2).name = "规范限值" & Sheets(gen_sheet).Cells(25, 7)
                    
                    With .SeriesCollection(2)
                        '选择线宽
                        .Format.Line.Weight = 2
                    
                        '选择线颜色
                        .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
                    
                        '选择线型-单线
                        .Format.Line.Style = msoLineSingle
                    
                        '选择线型-短线类型
                        .Format.Line.DashStyle = msoLineSolid   '实线
                        
                        '选择数据点类型
                        .MarkerStyle = 0
                        
                    End With
                    
            End If
        
        End If
        
        If name = "Y向剪重比" Then
        
            If Not IsEmpty(Sheets(gen_sheet).Cells(25, 7)) Then
            
                    '添加数据系列
                    .SeriesCollection.NewSeries
                    
                    '选择X轴系列值
                    .SeriesCollection(2).XValues = "={" & Sheets(gen_sheet).Cells(25, 7) & "," & Sheets(gen_sheet).Cells(24, 7) & "}"
                    
                    '选择Y轴系列值
                    .SeriesCollection(2).Values = "={0," & Num_all & "}"
                    
                    '选择系列标题
                    .SeriesCollection(2).name = "规范限值" & Sheets(gen_sheet).Cells(25, 7)
                    
                    With .SeriesCollection(2)
                        '选择线宽
                        .Format.Line.Weight = 2
                    
                        '选择线颜色
                        .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
                    
                        '选择线型-单线
                        .Format.Line.Style = msoLineSingle
                    
                        '选择线型-短线类型
                        .Format.Line.DashStyle = msoLineSolid   '实线
                        
                        '选择数据点类型
                        .MarkerStyle = 0
                        
                    End With
            
            End If
        
        End If
        
        '-------------------------------------------------------------刚度比限值
        If name = "X向刚度比" Then
        
            '添加数据系列
            .SeriesCollection.NewSeries
            
            '选择X轴系列值
            .SeriesCollection(2).XValues = "={1,1}"
            
            '选择Y轴系列值
            .SeriesCollection(2).Values = "={0," & Num_all & "}"
            
            '选择系列标题
            .SeriesCollection(2).name = "规范限值"
            
            With .SeriesCollection(2)
                '选择线宽
                .Format.Line.Weight = 2
            
                '选择线颜色
                .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
            
                '选择线型-短线类型
                .Format.Line.DashStyle = msoLineSolid   '实线
                
                '选择数据点类型
                .MarkerStyle = 0
                
            End With
        
        End If
        
        If name = "Y向刚度比" Then
        
            '添加数据系列
            .SeriesCollection.NewSeries
            
            '选择X轴系列值
            .SeriesCollection(2).XValues = "={1,1}"
            
            '选择Y轴系列值
            .SeriesCollection(2).Values = "={0," & Num_all & "}"
            
            '选择系列标题
            .SeriesCollection(2).name = "规范限值"
            
            With .SeriesCollection(2)
                '选择线宽
                .Format.Line.Weight = 2
            
                '选择线颜色
                .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
            
                '选择线型-短线类型
                .Format.Line.DashStyle = msoLineSolid   '实线
                
                '选择数据点类型
                .MarkerStyle = 0
                
            End With
        
        End If
        
        
        '-------------------------------------------------------------抗剪承载力比限值
        If name_X = "抗剪承载力比" Then
        
            '添加数据系列
            .SeriesCollection.NewSeries
            
            '选择X轴系列值
            .SeriesCollection(2).XValues = "={0.75,0.75}"
            
            '选择Y轴系列值
            .SeriesCollection(2).Values = "={0," & Num_all & "}"
            
            '选择系列标题
            .SeriesCollection(2).name = "规范限值"
            
            With .SeriesCollection(2)
                '选择线宽
                .Format.Line.Weight = 2
            
                '选择线颜色
                .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
            
                '选择线型-短线类型
                .Format.Line.DashStyle = msoLineSolid   '实线
                
                '选择数据点类型
                .MarkerStyle = 0
                
            End With
        
        End If
        
        
        '-------------------------------------------------------------质量比限值
        If name_X = "质量比" Then
        
            '添加数据系列
            .SeriesCollection.NewSeries
            
            '选择X轴系列值
            .SeriesCollection(2).XValues = "={1.5,1.5}"
            
            '选择Y轴系列值
            .SeriesCollection(2).Values = "={0," & Num_all & "}"
            
            '选择系列标题
            .SeriesCollection(2).name = "规范限值"
            
            With .SeriesCollection(2)
                '选择线宽
                .Format.Line.Weight = 2
            
                '选择线颜色
                .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            
                '选择线型-单线
                .Format.Line.Style = msoLineSingle
            
                '选择线型-短线类型
                .Format.Line.DashStyle = msoLineSolid   '实线
                
                '选择数据点类型
                .MarkerStyle = 0
                
            End With
        
        End If
        
        End With
        
        '.清空对象
        Set myChart = Nothing
            
    End With

    Debug.Print "绘图结束"

End Sub


