Attribute VB_Name = "模块_AddHeadline"
Option Explicit


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2015/4/19

'更新内容:
'1.ETABS新添数据列


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2015/4/17

'更新内容:
'1.ETABS工况名修改；


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2014/4/18

'更新内容:
'1.g表里添加底部框架柱倾覆力矩百分比；

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2014/3/4

'更新内容:
'1.增加ETABS工况名内容，现为 “EX(工况名)”格式

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/8/19 15:07

'更新内容:
'1.修改振型号处格式，不带小数


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/29 13:34

'更新内容:
'1.“最大位移角比”改为“最大层间位移比
'2. 增加计算日期栏

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/23 18:01

'更新内容:
'1.更改墙/柱表为可选变量,如调用中不写则不生成该表

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/19 10:09

'更新内容:
'1.更改位移比、刚度比、楼层抗剪承载力比单元格格式


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/18 17:01

'更新内容:
'1.添加层间位移角,最大位移比,层间位移比限值
'2.更改楼层，编号，层间位移角限值单元格格式


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/12 17:28

'更新内容:
'1.重新设计了general表格

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/7/3 17:28

'更新内容:
'1.调整了位移、层间位移角列的排序，以X和Y分开，方便取大值时包括风工况


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/5/27 19:17

'更新内容:
'1.dis工作表修改调整系数，增加轴压比等列
'2.gen工作表修改轴压比区域、最大位移比区域


'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/5/21 21:22

'更新内容:
'1.增加单位面积质量区域
'2.增加最大层间位移角出现的楼层、工况项位置

'///////////////////////////////////////////////////////////////////////////////////////////////////////

'更新时间: 2013/5/16 20:06

'///////////////////////////////////////////////////////////////////////////////////////////////////////


Sub AddHeadline(gen, dis, Optional column As String = "C", Optional wall As String = "W")


'计算运行时间
Dim sngStart As Single
sngStart = Timer



'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            工作表general的设定                       ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'======================================================================================================添加表格general的标题

Debug.Print "开始设定表格general的格式"
Debug.Print "……"

'清除工作表所有内容
Sheets(gen).Cells.Clear

'调整单元格宽高
Sheets(gen).Columns("A:A").ColumnWidth = 4
Sheets(gen).Columns("B:C").ColumnWidth = 10
Sheets(gen).Columns("D:G").ColumnWidth = 15
Sheets(gen).Rows("1:51").RowHeight = 13.5

'设为分页视图
Sheets(gen).Activate
ActiveWindow.View = xlPageBreakPreview
ActiveWindow.Zoom = 90

'加表格线
Call AddFormLine(gen, "b3:g54")

'加背景色
Call AddShadow(gen, "B3:C25", 6750105)
Call AddShadow(gen, "B27:C39", 6750105)
Call AddShadow(gen, "B41:C45", 6750105)
Call AddShadow(gen, "B47:C51", 6750105)
Call AddShadow(gen, "D5:D13", 6750105)
Call AddShadow(gen, "D15:D15", 6750105)
Call AddShadow(gen, "D17:D17", 6750105)
Call AddShadow(gen, "D19:D25", 6750105)
Call AddShadow(gen, "D39:D39", 6750105)
Call AddShadow(gen, "F4:F13", 6750105)
Call AddShadow(gen, "F15:F15", 6750105)
Call AddShadow(gen, "F17:F17", 6750105)
Call AddShadow(gen, "F19:F25", 6750105)
Call AddShadow(gen, "F38:F39", 6750105)
Call AddShadow(gen, "D27:G27", 6750105)
Call AddShadow(gen, "D41:G41", 6750105)
Call AddShadow(gen, "D47:G47", 6750105)
Call AddShadow(gen, "F14:F14", 6750105)
Call AddShadow(gen, "F16:F16", 6750105)
Call AddShadow(gen, "F18:F18", 6750105)
Call AddShadow(gen, "G14:G14", 5505023)
Call AddShadow(gen, "G16:G16", 5505023)
Call AddShadow(gen, "G18:G18", 5505023)
Call AddShadow(gen, "B53:D54", 6750105)
Call AddShadow(gen, "F53:F54", 6750105)

Debug.Print "设定表格general的格式完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"


Debug.Print "开始添加表格general的标题"
Debug.Print "……"

'------------------------------------------------------工作表general内的标题格式
With Sheets(gen)
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    .Cells.NumberFormatLocal = "0.00"
    '设置局部单元格特殊格式
    .range("G8:G9").NumberFormatLocal = "G/通用格式"
    .range("C28:C37").NumberFormatLocal = "G/通用格式"
    .Cells(4, 7).NumberFormatLocal = "yyyy 年 m 月 d 日"
    .Cells(14, 7).NumberFormatLocal = "# ???/???"
    .Cells(15, 7).NumberFormatLocal = "G/通用格式"
    .Cells(17, 7).NumberFormatLocal = "G/通用格式"
    .Cells(19, 7).NumberFormatLocal = "G/通用格式"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 4) = "计算结果记录表"
    .Cells(1, 4).Font.name = "黑体"
    .Cells(1, 4).Font.Size = "20"
    '合并单元格
    .range("D1:E2").MergeCells = True
    
    '-------------------------------------------------项目信息区
    '项目信息
    .Cells(3, 2) = "工程名称（路径）"
    .Cells(4, 6) = "计算日期"
    .Cells(4, 2) = "计算程序"
    .Cells(5, 2) = "计算参数"
    .Cells(5, 4) = "楼层自由度"
    .Cells(5, 6) = "周期折减系数"
    '合并单元格
    .range("B3:C3").MergeCells = True
    .range("B4:C4").MergeCells = True
    .range("B5:C5").MergeCells = True
    .range("D3:G3").MergeCells = True
    .range("D4:E4").MergeCells = True
    
    '-------------------------------------------------质量信息区
    
    '质量信息区标题
    .Cells(6, 2) = "质量"
    .Cells(6, 4) = "活载质量"
    .Cells(6, 6) = "附加质量"
    .Cells(7, 4) = "恒载质量"
    .Cells(7, 6) = "总质量"
    '合并单元格
    .range("B6:C7").MergeCells = True
    
    '-------------------------------------------------轴压比区
    
    '轴压比区标题
    .Cells(8, 2) = "最大轴压比"
    .Cells(8, 4) = "首层柱"
    .Cells(8, 6) = "编号"
    .Cells(9, 4) = "首层墙"
    .Cells(9, 6) = "编号"
    '合并单元格
    .range("B8:C9").MergeCells = True
    
    '-------------------------------------------------位移角区
    
    '位移角区标题
    .Cells(10, 2) = "层间位移角"
    .Cells(10, 3) = "风荷载"
    .Cells(10, 4) = "X向"
    .Cells(10, 6) = "Y向"
    .Cells(11, 3) = "地震"
    .Cells(11, 4) = "X向"
    .Cells(11, 6) = "Y向"
    .Cells(12, 4) = "X+5%向"
    .Cells(12, 6) = "Y+5%向"
    .Cells(13, 4) = "X-5%向"
    .Cells(13, 6) = "Y-5%向"
    .Cells(14, 2) = "最大层间位移角"
    .Cells(14, 6) = "限值"
    .Cells(15, 4) = "工况"
    .Cells(15, 6) = "楼层"
    .Cells(16, 2) = "最大位移比"
    .Cells(16, 6) = "限值"
    .Cells(17, 4) = "工况"
    .Cells(17, 6) = "楼层"
    .Cells(18, 2) = "最大层间位移比"
    .Cells(18, 6) = "限值"
    .Cells(19, 4) = "工况"
    .Cells(19, 6) = "楼层"
    '合并单元格
    .range("B10:B13").MergeCells = True
    .range("C11:C13").MergeCells = True
    .range("B14:C15").MergeCells = True
    .range("B16:C17").MergeCells = True
    .range("B18:C19").MergeCells = True
    .range("D14:E14").MergeCells = True
    .range("D16:E16").MergeCells = True
    .range("D18:E18").MergeCells = True
    
    '-------------------------------------------------刚重比区
    
    '轴压比区标题
    .Cells(20, 2) = "稳定性验算（刚重比）"
    .Cells(20, 4) = "X向"
    .Cells(20, 6) = "判断"
    .Cells(21, 4) = "Y向"
    .Cells(21, 6) = "判断"
    '合并单元格
    .range("B20:C21").MergeCells = True
    
    '-------------------------------------------------刚度比区
    
    '轴压比区标题
    .Cells(22, 2) = "最小刚度比"
    .Cells(22, 4) = "X向"
    .Cells(22, 6) = "Y向"
    '合并单元格
    .range("B22:C22").MergeCells = True
    
    '-------------------------------------------------楼层抗剪承载力比区
    
    '轴压比区标题
    .Cells(23, 2) = "层间受剪承载力比"
    .Cells(23, 4) = "X向"
    .Cells(23, 6) = "Y向"
    '合并单元格
    .range("B23:C23").MergeCells = True
    
    '-------------------------------------------------剪重比区
    
    '轴压比区标题
    .Cells(24, 2) = "最小剪重比"
    .Cells(24, 4) = "X向"
    .Cells(24, 6) = "X向限值"
    .Cells(25, 4) = "Y向"
    .Cells(25, 6) = "Y向限值"
    '合并单元格
    .range("B24:C25").MergeCells = True
    .range("B26:G26").MergeCells = True
    
    '-------------------------------------------------周期信息区
    
    '周期信息区标题
    .Cells(27, 2) = "动力特性"
    .Cells(27, 3) = "振型号"
    .Cells(28, 3) = "1"
    .Cells(29, 3) = "2"
    .Cells(30, 3) = "3"
    .Cells(31, 3) = "4"
    .Cells(32, 3) = "5"
    .Cells(33, 3) = "6"
    .Cells(34, 3) = "7"
    .Cells(35, 3) = "8"
    .Cells(36, 3) = "9"
    .Cells(37, 3) = "10"
    .Cells(38, 2) = "周期比"
    .Cells(38, 6) = "计算振型个数"
    .Cells(39, 2) = "振型参与质量系数"
    .Cells(27, 4) = "周期"
    .Cells(27, 5) = "转角"
    .Cells(27, 6) = "平动系数"
    .Cells(27, 7) = "扭转系数"
    .Cells(39, 4) = "X向"
    .Cells(39, 6) = "Y向"
    '合并单元格
    .range("B27:B37").MergeCells = True
    .range("B38:C38").MergeCells = True
    '.range("D38:E38").MergeCells = True
    .range("B39:C39").MergeCells = True
    .range("B40:G40").MergeCells = True
    
    '-------------------------------------------------结构抗力区
    
    '抗倾覆信息区标题
    .Cells(42, 2) = "风"
    .Cells(42, 3) = "X向"
    .Cells(43, 3) = "Y向"
    .Cells(44, 2) = "地震"
    .Cells(44, 3) = "X向"
    .Cells(45, 3) = "Y向"
    .Cells(41, 4) = "底层剪力"
    .Cells(41, 6) = "底层倾覆力矩"
    '合并单元格
    .range("B42:B43").MergeCells = True
    .range("B44:B45").MergeCells = True
    .range("D41:E41").MergeCells = True
    .range("F41:G41").MergeCells = True
    .range("D42:E42").MergeCells = True
    .range("F42:G42").MergeCells = True
    .range("D43:E43").MergeCells = True
    .range("F43:G43").MergeCells = True
    .range("D44:E44").MergeCells = True
    .range("F44:G44").MergeCells = True
    .range("D45:E45").MergeCells = True
    .range("F45:G45").MergeCells = True
    .range("B46:G46").MergeCells = True
    
    '-------------------------------------------------抗倾覆信息区
    
    '抗倾覆信息区标题
    .Cells(48, 2) = "风"
    .Cells(48, 3) = "X向"
    .Cells(49, 3) = "Y向"
    .Cells(50, 2) = "地震"
    .Cells(50, 3) = "X向"
    .Cells(51, 3) = "Y向"
    .Cells(47, 4) = "抗倾覆力矩Mr"
    .Cells(47, 5) = "倾覆力矩Mov"
    .Cells(47, 6) = "比值Mr/Mov"
    .Cells(47, 7) = "零应力区(%)"
    .range("B48:B49").MergeCells = True
    .range("B50:B51").MergeCells = True
    
    
    '-------------------------------------------------框架柱及短肢墙地震倾覆力矩百分比区
    
    '框架柱及短肢墙地震倾覆力矩百分比标题
    .range("B52:G52").MergeCells = True
    .Cells(53, 2) = "框架柱及短肢墙地震倾覆力矩百分比"
    .range("B53:C54").MergeCells = True
    
    .Cells(53, 4) = "X向"
    .Cells(54, 4) = "Y向"
    
    .Cells(53, 6) = "X向"
    .Cells(54, 6) = "Y向"
    
End With

Debug.Print "添加表格general的标题完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"



'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                          工作表distribution的设定                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************




'======================================================================================================设定表格distribution的格式


Debug.Print "开始设定表格distribution的格式"
Debug.Print "……"

Sheets(dis).Cells.Clear

'加表格线
Call AddFormLine(dis, "A1:BG200")

'加背景色
'层号区
Call AddShadow(dis, "A1:A2", 10092441)
'刚度比区
Call AddShadow(dis, "B1:C2", 6750207)
'刚度区
Call AddShadow(dis, "D1:E2", 10092441)
'风荷载区
Call AddShadow(dis, "F1:I2", 6750207)
'地震荷载区
Call AddShadow(dis, "J1:Q2", 10092441)
'位移区
Call AddShadow(dis, "R1:Y2", 6750207)
'层间位移角区
Call AddShadow(dis, "Z1:AG2", 10092441)
'位移比区
Call AddShadow(dis, "AH1:AM2", 6750207)
'层间位移比区
Call AddShadow(dis, "AN1:AS2", 10092441)
'抗剪承载力区
Call AddShadow(dis, "AT1:AU2", 6750207)
'调整剪力区
Call AddShadow(dis, "AV1:BA2", 10092441)
'质量分布区
Call AddShadow(dis, "BB1:BC2", 6750207)
'柱最大轴压比区
Call AddShadow(dis, "BD1:BE2", 10092441)
'墙最大轴压比区
Call AddShadow(dis, "BF1:BG2", 6750207)

'冻结标题
Sheets(dis).Select
range("B3").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 2
End With
ActiveWindow.FreezePanes = True


Debug.Print "设定表格distribution的格式完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"

'======================================================================================================添加表格distribution的标题


'-----------------------------------------------------表格distribution的标题格式


Debug.Print "开始添加表格distribution的标题"
Debug.Print "……"


With Sheets(dis)
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------层号区
    
   .Cells(1, 1) = "层号"
    .range("A1:A2").MergeCells = True
    
    '-------------------------------------------------刚度比区
    
    .Cells(1, 2) = "刚度比"
    .Cells(2, 2) = "Ratx"
    .Cells(2, 3) = "Raty"
    .range("B1:C1").MergeCells = True
    
    '-------------------------------------------------刚度区
    
    .Cells(1, 4) = "刚度"
    .Cells(2, 4) = "RJX3"
    .Cells(2, 5) = "RJY3"
    .range("D1:E1").MergeCells = True
    
    '-------------------------------------------------风荷载区
    
    .Cells(1, 6) = "风荷载"
    .Cells(2, 6) = "剪力X"
    .Cells(2, 7) = "弯矩X"
    .Cells(2, 8) = "剪力Y"
    .Cells(2, 9) = "弯矩Y"
    .range("F1:I1").MergeCells = True
    
    '-------------------------------------------------地震荷载区
    
    .Cells(1, 10) = "地震荷载"
    .Cells(2, 10) = "剪力X"
    .Cells(2, 11) = "弯矩X"
    .Cells(2, 12) = "剪重比X"
    .Cells(2, 13) = "调整后剪重比X"
    .Cells(2, 14) = "剪力Y"
    .Cells(2, 15) = "弯矩Y"
    .Cells(2, 16) = "剪重比Y"
    .Cells(2, 17) = "调整后剪重比Y"
    .range("J1:Q1").MergeCells = True
    
    '-------------------------------------------------位移区
    
    .Cells(1, 18) = "位移"
    .Cells(2, 18) = "EX"
    .Cells(2, 19) = "EX +"
    .Cells(2, 20) = "EX -"
    .Cells(2, 21) = "Wind X"
    .Cells(2, 22) = "EY  "
    .Cells(2, 23) = "EY +"
    .Cells(2, 24) = "EY -"
    .Cells(2, 25) = "Wind Y"
    .range("R1:Y1").MergeCells = True
    
    '-------------------------------------------------层间位移角区
    
    .Cells(1, 26) = "层间位移角"
    .Cells(2, 26) = "EX"
    .Cells(2, 27) = "EX +"
    .Cells(2, 28) = "EX -"
    .Cells(2, 29) = "Wind X"
    .Cells(2, 30) = "EY  "
    .Cells(2, 31) = "EY +"
    .Cells(2, 32) = "EY -"
    .Cells(2, 33) = "Wind Y"
    .range("Z1:AG1").MergeCells = True
        
    '-------------------------------------------------位移比区
    
    .Cells(1, 34) = "位移比"
    .Cells(2, 34) = "EX  "
    .Cells(2, 35) = "EX +"
    .Cells(2, 36) = "EX -"
    .Cells(2, 37) = "EY  "
    .Cells(2, 38) = "EY +"
    .Cells(2, 39) = "EY -"
    .range("AH1:AM1").MergeCells = True
    
    '-------------------------------------------------层间位移比区
    
    .Cells(1, 40) = "层间位移比"
    .Cells(2, 40) = "EX  "
    .Cells(2, 41) = "EX +"
    .Cells(2, 42) = "EX -"
    .Cells(2, 43) = "EY  "
    .Cells(2, 44) = "EY +"
    .Cells(2, 45) = "EY -"
    .range("AN1:AS1").MergeCells = True
    
    '-------------------------------------------------抗剪承载力区
    
    .Cells(1, 46) = "楼层抗剪承载力"
    .Cells(2, 46) = "Ratio_X"
    .Cells(2, 47) = "Ratio_Y"
    .range("AT1:AU1").MergeCells = True
    
    '-------------------------------------------------调整剪力区
    
    .Cells(1, 48) = "调整后的剪力"
    .Cells(2, 48) = "V_Col_X"
    .Cells(2, 49) = "Ratio_VX"
    .Cells(2, 50) = "Cof_VX"
    .Cells(2, 51) = "V_Col_Y"
    .Cells(2, 52) = "Ratio_VY"
    .Cells(2, 53) = "Cof_VY"
    .range("AV1:BA1").MergeCells = True
    
    '-------------------------------------------------质量分布区
    
    .Cells(1, 54) = "单位面积质量分布"
    .Cells(2, 54) = "单位面积质量"
    .Cells(2, 55) = "质量比"
    .range("BB1:BC1").MergeCells = True
    
    '-------------------------------------------------轴压比区
    
    .Cells(1, 56) = "柱最大轴压比"
    .Cells(2, 56) = "Uc_Max"
    .Cells(2, 57) = "编号"
    .range("BD1:BE1").MergeCells = True
    
    .Cells(1, 58) = "墙最大轴压比"
    .Cells(2, 58) = "Uc_Max"
    .Cells(2, 59) = "编号"
    .range("BF1:BG1").MergeCells = True
    
    
    '-------------------------------------------------ETABS匹配内容名称与荷载工况
    If dis = "d_E" Then
      '正向地震反应谱相关
      If Not IsEmpty(OUTReader_Main.ComboBox_SPEC_X) Then
        '刚度比
        .Cells(2, 2) = .Cells(2, 2) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '刚度
        .Cells(2, 4) = .Cells(2, 4) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '剪力
        .Cells(2, 10) = .Cells(2, 10) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '弯矩，规定水平地震力
        .Cells(2, 11) = .Cells(2, 11) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '剪重比
        .Cells(2, 12) = .Cells(2, 12) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '调整后剪重比
        .Cells(2, 13) = .Cells(2, 13) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '位移，规定水平地震力
        .Cells(2, 18) = .Cells(2, 18) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '层间位移角
        .Cells(2, 26) = .Cells(2, 26) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 34) = .Cells(2, 34) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 40) = .Cells(2, 40) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        '框架剪力弯矩占比
        .Cells(2, 76) = "Co_V" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 77) = "Co_V" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        .Cells(2, 78) = "Wa_V" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 79) = "Wa_V" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        .Cells(2, 80) = "V" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 81) = "V" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        
        .Cells(2, 82) = "Co_M" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 83) = "Co_M" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        .Cells(2, 84) = "Wa_M" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 85) = "Wa_M" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        .Cells(2, 86) = "M" & "(" & OUTReader_Main.ComboBox_SPEC_X.Text & ")"
        .Cells(2, 87) = "M" & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        
      End If
      
      If Not IsEmpty(OUTReader_Main.ComboBox_SPEC_Y) Then
        '刚度比
        .Cells(2, 3) = .Cells(2, 3) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '刚度
        .Cells(2, 5) = .Cells(2, 5) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '剪力
        .Cells(2, 14) = .Cells(2, 14) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '弯矩，规定水平地震力
        .Cells(2, 15) = .Cells(2, 15) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '剪重比
        .Cells(2, 16) = .Cells(2, 16) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '调整后剪重比
        .Cells(2, 17) = .Cells(2, 17) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '位移，规定水平地震力
        .Cells(2, 22) = .Cells(2, 22) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '层间位移角
        .Cells(2, 30) = .Cells(2, 30) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 37) = .Cells(2, 37) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 43) = .Cells(2, 43) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_Y.Text & ")"
        
      End If
      
      '偏心地震反应谱相关
      If Not IsEmpty(OUTReader_Main.ComboBox_SPEC_XEcc) Then
        '正向偏心地震反应谱相关
        '位移，规定水平地震力
        .Cells(2, 19) = .Cells(2, 19) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & ")"
        '层间位移角
        .Cells(2, 27) = .Cells(2, 27) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 35) = .Cells(2, 35) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 41) = .Cells(2, 41) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & ")"
        
        '负向偏心地震反应谱相关
        '位移，规定水平地震力
        .Cells(2, 20) = .Cells(2, 20) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & ")"
        '层间位移角
        .Cells(2, 28) = .Cells(2, 28) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 36) = .Cells(2, 36) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 42) = .Cells(2, 42) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & ")"
        
         
         
         .Cells(2, 70) = "Disp_Max" & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & ")"
         .Cells(2, 71) = "Disp_Max" & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & ")"
         
      End If
      
      If Not IsEmpty(OUTReader_Main.ComboBox_SPEC_YEcc) Then
        '正向偏心地震反应谱相关
        '位移，规定水平地震力
        .Cells(2, 23) = .Cells(2, 23) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & ")"
        '层间位移角
        .Cells(2, 31) = .Cells(2, 31) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 38) = .Cells(2, 38) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 44) = .Cells(2, 44) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & ")"
        
        '负向偏心地震反应谱相关
        '位移，规定水平地震力
        .Cells(2, 24) = .Cells(2, 24) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & ")"
        '层间位移角2
        .Cells(2, 32) = .Cells(2, 32) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & ")"
        '位移比，规定水平地震力
        .Cells(2, 39) = .Cells(2, 39) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & ")"
        '层间位移比，规定水平地震力
        .Cells(2, 45) = .Cells(2, 45) & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & ")"
        
         .Cells(2, 72) = "Disp_Max" & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & ")"
         .Cells(2, 73) = "Disp_Max" & vbCrLf & "(" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & ")"
      End If
      
      '风荷载相关
      If Not IsEmpty(OUTReader_Main.ComboBox_Wind_X) Then
        '剪力
        .Cells(2, 6) = .Cells(2, 6) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_X.Text & ")"
        '弯矩
        .Cells(2, 7) = .Cells(2, 7) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_X.Text & ")"
        '位移
        .Cells(2, 21) = .Cells(2, 21) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_X.Text & ")"
        '层间位移角
        .Cells(2, 29) = .Cells(2, 29) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_X.Text & ")"
        
      End If
      
      If Not IsEmpty(OUTReader_Main.ComboBox_Wind_Y) Then
        '剪力
        .Cells(2, 8) = .Cells(2, 8) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_Y.Text & ")"
        '弯矩
        .Cells(2, 9) = .Cells(2, 9) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_Y.Text & ")"
        '位移
        .Cells(2, 25) = .Cells(2, 25) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_Y.Text & ")"
        '层间位移角
        .Cells(2, 33) = .Cells(2, 33) & vbCrLf & "(" & OUTReader_Main.ComboBox_Wind_Y.Text & ")"
        
      End If
      
      
      
    End If

    
    
End With


Debug.Print "添加表格distribution的标题结束"
Debug.Print "……"


'单元格宽度自动调整
Sheets(dis).Cells.EntireColumn.AutoFit
'设置小数点后位数
Sheets(dis).Cells.NumberFormatLocal = "G/通用格式"
Sheets(dis).Columns("B:C").NumberFormatLocal = "0.00"
Sheets(dis).Columns("AT:AU").NumberFormatLocal = "0.00"
Sheets(dis).Columns("AH:AS").NumberFormatLocal = "0.00"

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                          工作表Column及Wall的设定                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

Debug.Print "开始设定表格Column及Wall的格式"
Debug.Print "……"

'--------------------------------------------------------------------------Column
Debug.Print column
If column <> "C" Then
Sheets(column).Cells.Clear
With Sheets(column)
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '列宽
    .Columns.ColumnWidth = 5.5
End With
'构件编号
Sheets(column).Cells(1, 1) = "编号"
'冻结首行首列
Sheets(column).Select
range("B2").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
End If

'--------------------------------------------------------------------------Wall
If wall <> "W" Then
Sheets(wall).Cells.Clear
With Sheets(wall)
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '列宽
    .Columns.ColumnWidth = 5.5
End With
'构件编号
Sheets(wall).Cells(1, 1) = "编号"
'冻结首行首列
Sheets(wall).Select
range("B2").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True

Debug.Print "添加表格Column及Wall的标题结束"
Debug.Print "……"
End If

End Sub


'添加区域表格线模块

Sub AddFormLine(shname, sel)

With Sheets(shname).range(sel)
    '定义区域内单元格水平线
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    '定义区域内单元格竖直线
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    '定义区域左侧线
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    '定义区域右侧线
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    '定义区域上侧线
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    '定义区域下侧线
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
End With

End Sub

'添加单元格底色模块

Sub AddShadow(shname, sel, color)

With Sheets((shname)).range(sel).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    '变量color为自定义的颜色代码
    .color = color
    '颜色明度，-1（最暗）到 1（最亮）
    .TintAndShade = 0.3
    .PatternTintAndShade = 0
End With
    
End Sub

