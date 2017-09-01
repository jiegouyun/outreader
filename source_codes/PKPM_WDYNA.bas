Attribute VB_Name = "PKPM_WDYNA"
Option Explicit


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            PKPM_WDYNA.OUT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/11/21
'1.匹配PKPM v2.2版。新版本名称是 v2.2中文版，不知道有没有v2.2英文版，故条件判断
'包含了旧版英文版和新版中文版（中文版判断条件在后），以后如无需要，可把英文版判断语句删除

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/8
'1.增加表头背景色

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/29
'1.添加正负35%，正负20%反应谱剪力曲线

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/9/10
'1.修改了位移角读取时的关键词,改为"1"


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27
'1.调整反应谱汇总数据位置到末尾，方便绘图，添加纪录时程波总数。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/19
'1.修改读取反应谱Y向位移角bug；
'2.出去地震波名称的中括号；
'3.将地震波对应列位置合并；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/25

'更新内容：
'1.增加反应谱数据判断和提示代码，防止没有反应数据时时程分析数据相除报错。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/26

'更新内容：
'1.删去平均值/反应谱，加上位移角；
'2.将反应谱数据移入，以方便后期绘图；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/16 21:30


Sub OUTReader_PKPM_WDYNA(Path)

'计算运行时间
Dim sngStart As Single
sngStart = Timer

'======================================================================================================提示反应谱数据
'If Sheets("d_P").Cells(3, 6) = "" Or Sheets("d_P").Cells(3, 10) = "" Or Sheets("d_P").Cells(3, 18) = "" Then
   ' MsgBox "缺少反应谱数据，请返回重新选择。"
'Else


'======================================================================================================设定表格Elastic-Dynamic的格式
'定义结果表格名称
Dim ela As String
ela = "e_P"

Debug.Print "开始设定表格Elastic-Dynamic的格式"
Debug.Print "……"

'清除工作表所有内容
Sheets(ela).Cells.Clear

Debug.Print "设定表格Elastic-Dynamic的格式完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"

'======================================================================================================添加表格Elastic-Dynamic的标题

Debug.Print "开始添加表格Elastic-Dynamic的标题"
Debug.Print "……"

'------------------------------------------------------工作表Elastic-Dynamic内的标题格式
With Sheets(ela)
    
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
    
    '-------------------------------------------------汇总表格区1
    
    '项目信息区标题
    .Cells(2, 1) = "时程波数"
    .Cells(4, 1) = "作用工况"
    .Cells(4, 2) = "作用方向=0°"
    .Cells(4, 5) = "作用方向=90°"
    .Cells(5, 2) = "基底剪力"
    .Cells(5, 3) = "时程/反应谱"
    .Cells(5, 4) = "位移角"
    .Cells(5, 5) = "基底剪力"
    .Cells(5, 6) = "时程/反应谱"
    .Cells(5, 7) = "位移角"
    .range("A4:A5").MergeCells = True
    .range("B4:D4").MergeCells = True
    .range("E4:G4").MergeCells = True
    
    '-------------------------------------------------汇总表格区2
    
    '项目信息区标题
    '.Cells(18, 1) = "作用工况"
    '.Cells(18, 2) = "作用方向=0°"
    '.Cells(18, 5) = "作用方向=90°"
    '.Cells(19, 2) = "位移角"
    '.Cells(19, 3) = "时程/反应谱"
    '.Cells(19, 4) = "平均值/反应谱"
    '.Cells(19, 5) = "基底剪力"
    '.Cells(19, 6) = "时程/反应谱"
    '.Cells(19, 7) = "平均值/反应谱"
    '.Cells(20, 1) = "反应谱"
    '.range("A18:A19").MergeCells = True
    '.range("B18:D18").MergeCells = True
    '.range("E18:G18").MergeCells = True
    
End With

'加表格线
Call AddFormLine(ela, "A1:DZ200")

Debug.Print "添加表格Elastic-Dynamic的标题完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring As String

'path = Sheets("debug").Cells(2, 2).Value

'定义data为读入行的字符串
Dim data As String

Dim i, j, n As Integer

'定义m记录地震波序号
Dim m As Integer
m = 0

'定义颜色变量
Dim Temp_Colour, Colour As Long
Temp_Colour = 1
'=============================================================================================================================生成文件读取路径

'指定文件名为wdyna.out
Filename = "WDYNA.OUT"

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
'Debug.Print path
'Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================第一次循环逐行读取文本

Debug.Print "开始第一次遍历结果文件wdyna.out"
Debug.Print "……"

Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    '记录行数
    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------读取最大加速度
    '加一个Mid的简单判断，不需要每一句都运行正则判断
    '备注：在用正则判断并提取最大加速度、地震波名称前，加一个MID语句删选非关键行，运行时间缩短为原来的40%左右
    '主方向
    If Mid(data, 46, 5) = "AmaxX" Or Mid(data, 23, 5) = "AmaxX" Then
        Sheets(ela).Cells(3, 1).Value = "AmaxX="
        Sheets(ela).Cells(3, 2).Value = extractNumberFromString(data, 1)
        Debug.Print "读取主方向最大加速度"
        Debug.Print "……"
    End If
    
    '次方向
    If Mid(data, 46, 5) = "AmaxY" Or Mid(data, 23, 5) = "AmaxY" Then
        Sheets(ela).Cells(3, 3).Value = "AmaxY="
        Sheets(ela).Cells(3, 4).Value = extractNumberFromString(data, 1)
        Debug.Print "读取次方向最大加速度"
        Debug.Print "……"
    End If
    
    '竖直方向
    If Mid(data, 46, 5) = "AmaxZ" Or Mid(data, 23, 5) = "AmaxZ" Then
        Sheets(ela).Cells(3, 5).Value = "AmaxZ="
        Sheets(ela).Cells(3, 6).Value = extractNumberFromString(data, 1)
        Debug.Print "读取竖直方向最大加速度"
        Debug.Print "……"
    End If
    
    
    '-------------------------------------------------------------------------------------------读取地震波名称
    
    '加一个Mid的简单判断，如果符合再提取地震波名称，这样能提高整体效率
    If Mid(data, 2, 52) = "==========The Maximum Response of the Seismic Wave [" Or Mid(data, 2, 20) = "================== [" Then
    
        '地震波名称格式为"[名称]"
'        If CheckRegExpfromString(data, "\[.+\]") Then
            '记录地震波的序号
            m = m + 1
            '记录地震波总数
            Sheets(ela).Cells(2, 2) = m
            '将名称写入汇总表
            Dim name As String
            name = StringfromStringforReg(data, "\[.+\]", 1)
            name = Replace(name, "[", "")
            name = Replace(name, "]", "")
            Sheets(ela).Cells(m + 5, 1) = name
            '将标题项写入层分布表
            Sheets(ela).Cells(1, 10 + (m - 1) * 6) = name
            Sheets(ela).Cells(2, (m - 1) * 6 + 10) = "层间位移角"
            Sheets(ela).Cells(2, (m - 1) * 6 + 13) = "层间位移角"
            Sheets(ela).Cells(2, (m - 1) * 6 + 11) = "剪力"
            Sheets(ela).Cells(2, (m - 1) * 6 + 14) = "剪力"
            Sheets(ela).Cells(2, (m - 1) * 6 + 12) = "倾覆弯矩"
            Sheets(ela).Cells(2, (m - 1) * 6 + 15) = "倾覆弯矩"
            
            '加背景色
            If Temp_Colour > 0 Then
              Colour = 10091441
            Else
              Colour = 6750207
            End If
    
            Sheets(ela).range(Cells(1, (m - 1) * 6 + 10), Cells(2, (m - 1) * 6 + 15)).Interior.color = Colour
            Temp_Colour = -1 * Temp_Colour
            
            Debug.Print "第" & m & "条地震波名称读取完毕"
            Debug.Print "运行时间: " & Timer - sngStart
            Debug.Print "……"
    
'        End If
    End If
    
    '所有地震波结果完了之后是两组数据，一组是最大值的平均值
    If Mid(data, 12, 50) = "The Average of Max_Response of These Seismic Waves" Or Mid(data, 28, 11) = "所有地震波反应的平均值" Then
        '记录地震波的序号
        m = m + 1
        '将名称写入汇总表
        Sheets(ela).Cells(m + 5, 1) = "平均值"
        '将标题项写入层分布表
        Sheets(ela).Cells(1, 10 + (m - 1) * 6) = "平均值"
        Sheets(ela).Cells(2, (m - 1) * 6 + 10) = "层间位移角"
        Sheets(ela).Cells(2, (m - 1) * 6 + 13) = "层间位移角"
        Sheets(ela).Cells(2, (m - 1) * 6 + 11) = "剪力"
        Sheets(ela).Cells(2, (m - 1) * 6 + 14) = "剪力"
        Sheets(ela).Cells(2, (m - 1) * 6 + 12) = "倾覆弯矩"
        Sheets(ela).Cells(2, (m - 1) * 6 + 15) = "倾覆弯矩"
        
        '加背景色
        If Temp_Colour > 0 Then
          Colour = 10091441
        Else
          Colour = 6750207
        End If

        Sheets(ela).range(Cells(1, (m - 1) * 6 + 10), Cells(2, (m - 1) * 6 + 15)).Interior.color = Colour
        Temp_Colour = -1 * Temp_Colour
        
        Debug.Print "The Average of Max_Response名称读取完毕"
        Debug.Print "运行时间: " & Timer - sngStart
        Debug.Print "……"
        
    End If
    
    '另一组是最大反应的最大值
    If Mid(data, 12, 50) = "The Maximum of Max_Response of These Seismic Waves" Or Mid(data, 28, 11) = "所有地震波反应的包络值" Then
        '记录地震波的序号
        m = m + 1
        '将名称写入汇总表
        Sheets(ela).Cells(m + 5, 1) = "最大值"
        '将标题项写入层分布表
        Sheets(ela).Cells(1, 10 + (m - 1) * 6) = "最大值"
        Sheets(ela).Cells(2, (m - 1) * 6 + 10) = "层间位移角"
        Sheets(ela).Cells(2, (m - 1) * 6 + 13) = "层间位移角"
        Sheets(ela).Cells(2, (m - 1) * 6 + 11) = "剪力"
        Sheets(ela).Cells(2, (m - 1) * 6 + 14) = "剪力"
        Sheets(ela).Cells(2, (m - 1) * 6 + 12) = "倾覆弯矩"
        Sheets(ela).Cells(2, (m - 1) * 6 + 15) = "倾覆弯矩"
        
        '加背景色
        If Temp_Colour > 0 Then
          Colour = 10091441
        Else
          Colour = 6750207
        End If

        Sheets(ela).range(Cells(1, (m - 1) * 6 + 10), Cells(2, (m - 1) * 6 + 15)).Interior.color = Colour
        Temp_Colour = -1 * Temp_Colour
        
        Debug.Print "The Maximum of Max_Response名称读取完毕"
        Debug.Print "运行时间: " & Timer - sngStart
        Debug.Print "……"
        
    End If
    
Loop

Close #i

Sheets(ela).Cells(m + 6, 1) = "反应谱"
Sheets(ela).Cells(m + 7, 1) = "65%反应谱"
Sheets(ela).Cells(m + 8, 1) = "135%反应谱"
Sheets(ela).Cells(m + 9, 1) = "80%反应谱"
Sheets(ela).Cells(m + 10, 1) = "120%反应谱"

Debug.Print "第一次循环完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"

'=============================================================================================================================第二次循环逐行读取文本

'由于本文本格式的特殊性，嵌套循环时，难以同时兼顾地震波名称读取和数据存取(无法找到合适的退出小循环的分隔符)，在一次循环里全部读取较困难
'本次循环的判断依据以地震的作用角度（如"Angle :     0.000"）为关键词，依次读取每一条地震波在不同作用角度下主方向的结果数据

Debug.Print "开始第二次遍历结果文件wdyna.out"
Debug.Print "……"


'初始化m、n值
m = 0
n = 0

'打开结果文件
Open (filepath) For Input Access Read As #i

Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    '记录行数
    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------读取地震作用角度=0°结果
    
    If Mid(data, 42, 17) = "Angle :     0.000" Or Mid(data, 6, 20) = "主方向作用方向角 :     0.000" Then
        
        m = m + 1
        Sheets(ela).range(Cells(1, (m - 1) * 6 + 10), Cells(1, (m - 1) * 6 + 15)).MergeCells = True
        '---------------------------------------------------------------------------------------嵌套第一个循环
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "Jmax") = True Then
                '-------------------------------------------------------------------------------嵌套第二个循环
                Do While Not EOF(1)
                Line Input #i, data
'                有塔楼号则为数据提取行
                    If StringfromStringforReg(data, "\S+", 2) = "1" Then
'                        结果文件中数据对应层号从大至小，统一为从小到大排列
'                        j为读取行数据写入表格的行数，跳过两行标题行
                        j = extractNumberFromString(data, 1) + 2
'                        逐一写入最大层间位移角
                        Sheets(ela).Cells(j, (m - 1) * 6 + 10).Value = StringfromStringforReg(data, "\b\d+\.\s", 1)
                    End If
'                    因为位移角和剪力弯矩不在同一列表内，检测到Shear，退出第二个嵌套循环
                    If CheckRegExpfromString(data, "Shear") Then
                        Exit Do
                    End If
                Loop
            End If
            
            If CheckRegExpfromString(data, "(kN)") = True Then
'                -------------------------------------------------------------------------------嵌套第二个循环
                Do While Not EOF(1)
                    Line Input #i, data
'                    有塔楼号则为数据提取行
                    If StringfromStringforReg(data, "\S+", 2) = "1" Then
'                        结果文件中数据对应层号从大至小，统一为从小到大排列
'                        j为读取行数据写入表格的行数，跳过两行标题行
                        j = extractNumberFromString(data, 1) + 2
'                        逐一写入剪力、弯矩、楼层
                        Sheets(ela).Cells(j, 11 + (m - 1) * 6) = StringfromStringforReg(data, "\S+", 3)
                        Sheets(ela).Cells(j, 12 + (m - 1) * 6) = StringfromStringforReg(data, "\S+", 5)
                        Sheets(ela).Cells(j, 9) = StringfromStringforReg(data, "\S+", 1)
                    End If
                    
                    If Mid(data, 40, 15) = "Minor-Direction" Or Mid(data, 6, 4) = "次方向：" Then
                        Exit Do
                    End If
                Loop
                Exit Do
            End If
      
        Loop
        
    End If

    '-------------------------------------------------------------------------------------------读取地震作用角度=90°结果
    If Mid(data, 42, 17) = "Angle :    90.000" Or Mid(data, 6, 20) = "主方向作用方向角 :    90.000" Then

        '---------------------------------------------------------------------------------------嵌套第一个循环
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "Jmax") = True Then
                '-------------------------------------------------------------------------------嵌套第二个循环
                Do While Not EOF(1)
                Line Input #i, data
'                有塔楼号则为数据提取行
                    If StringfromStringforReg(data, "\S+", 2) = "1" Then
'                        结果文件中数据对应层号从大至小，统一为从小到大排列
'                        j为读取行数据写入表格的行数，跳过两行标题行
                        j = extractNumberFromString(data, 1) + 2
'                        逐一写入最大层间位移角
                        Sheets(ela).Cells(j, (m - 1) * 6 + 13).Value = StringfromStringforReg(data, "\b\d+\.\s", 1)
                    End If
'                    因为位移角和剪力弯矩不在同一列表内，检测到Shear，退出第二个嵌套循环
                    If CheckRegExpfromString(data, "Shear") Then
                        Exit Do
                    End If
                Loop
            End If
            
            If CheckRegExpfromString(data, "(kN)") = True Then
'                -------------------------------------------------------------------------------嵌套第二个循环
                Do While Not EOF(1)
                    Line Input #i, data
'                    有塔楼号则为数据提取行
                    If StringfromStringforReg(data, "\S+", 2) = "1" Then
'                        结果文件中数据对应层号从大至小，统一为从小到大排列
'                        j为读取行数据写入表格的行数，跳过两行标题行
                        j = extractNumberFromString(data, 1) + 2
'                        逐一写入剪力、弯矩
                        Sheets(ela).Cells(j, 14 + (m - 1) * 6) = StringfromStringforReg(data, "\S+", 3)
                        Sheets(ela).Cells(j, 15 + (m - 1) * 6) = StringfromStringforReg(data, "\S+", 5)
'                        Sheets(ela).Cells(j, 9) = StringfromStringforReg(data, "\S+", 1)
                    End If
                    
                    If Mid(data, 40, 15) = "Minor-Direction" Or Mid(data, 6, 4) = "次方向：" Then
                        Exit Do
                    End If
                Loop
                Exit Do
            End If
            
        Loop
        
    End If
    
Loop

Close #i

Debug.Print "第二次循环完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"


'=============================================================================================================================将反应谱数据转入
Sheets(ela).Cells(1, 10 + m * 6) = "反应谱"
Sheets(ela).range(Cells(1, m * 6 + 10), Cells(1, m * 6 + 23)).MergeCells = True
Sheets(ela).Cells(2, m * 6 + 10) = "X层间位移角"
Sheets(ela).Cells(2, m * 6 + 13) = "Y层间位移角"
Sheets(ela).Cells(2, m * 6 + 11) = "X剪力"
Sheets(ela).Cells(2, m * 6 + 14) = "Y剪力"
Sheets(ela).Cells(2, m * 6 + 12) = "X倾覆弯矩"
Sheets(ela).Cells(2, m * 6 + 15) = "Y倾覆弯矩"
Sheets(ela).Cells(2, m * 6 + 16) = "65%X剪力"
Sheets(ela).Cells(2, m * 6 + 17) = "135%X剪力"
Sheets(ela).Cells(2, m * 6 + 18) = "80%X剪力"
Sheets(ela).Cells(2, m * 6 + 19) = "120%X剪力"
Sheets(ela).Cells(2, m * 6 + 20) = "65%Y剪力"
Sheets(ela).Cells(2, m * 6 + 21) = "135%Y剪力"
Sheets(ela).Cells(2, m * 6 + 22) = "80%Y剪力"
Sheets(ela).Cells(2, m * 6 + 23) = "120%Y剪力"

'加背景色
If Temp_Colour > 0 Then
  Colour = 10091441
Else
  Colour = 6750207
End If

Sheets(ela).range(Cells(1, m * 6 + 10), Cells(2, m * 6 + 23)).Interior.color = Colour
Temp_Colour = -1 * Temp_Colour

'确定楼层数
Dim NN As Integer: NN = Cells(Rows.Count, "j").End(3).Row - 2

'层间位移角
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 10), Sheets(ela).Cells(NN + 2, m * 6 + 10)).Value = Sheets("d_P").range("Z3:" & "Z" & NN + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 13), Sheets(ela).Cells(NN + 2, m * 6 + 13)).Value = Sheets("d_P").range("AD3:" & "AD" & NN + 2).Value
'剪力
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 11), Sheets(ela).Cells(NN + 2, m * 6 + 11)).Value = Sheets("d_P").range("J3:" & "J" & NN + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 14), Sheets(ela).Cells(NN + 2, m * 6 + 14)).Value = Sheets("d_P").range("N3:" & "N" & NN + 2).Value
'弯矩
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 12), Sheets(ela).Cells(NN + 2, m * 6 + 12)).Value = Sheets("d_P").range("K3:" & "K" & NN + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, m * 6 + 15), Sheets(ela).Cells(NN + 2, m * 6 + 15)).Value = Sheets("d_P").range("O3:" & "O" & NN + 2).Value

For i = 1 To NN
'X正负35%剪力
Sheets(ela).Cells(i + 2, m * 6 + 16) = Sheets(ela).Cells(i + 2, m * 6 + 11) * 0.65
Sheets(ela).Cells(i + 2, m * 6 + 17) = Sheets(ela).Cells(i + 2, m * 6 + 11) * 1.35
'X正负20%剪力
Sheets(ela).Cells(i + 2, m * 6 + 18) = Sheets(ela).Cells(i + 2, m * 6 + 11) * 0.8
Sheets(ela).Cells(i + 2, m * 6 + 19) = Sheets(ela).Cells(i + 2, m * 6 + 11) * 1.2
'Y正负35%剪力
Sheets(ela).Cells(i + 2, m * 6 + 20) = Sheets(ela).Cells(i + 2, m * 6 + 14) * 0.65
Sheets(ela).Cells(i + 2, m * 6 + 21) = Sheets(ela).Cells(i + 2, m * 6 + 14) * 1.35
'Y正负20%剪力
Sheets(ela).Cells(i + 2, m * 6 + 22) = Sheets(ela).Cells(i + 2, m * 6 + 14) * 0.8
Sheets(ela).Cells(i + 2, m * 6 + 23) = Sheets(ela).Cells(i + 2, m * 6 + 14) * 1.2
Next i

'=============================================================================================================================填写汇总表格

'读取反应谱的基底剪力
Sheets(ela).Cells(m + 6, 2) = Sheets("d_P").Cells(3, 10)
Sheets(ela).Cells(m + 6, 5) = Sheets("d_P").Cells(3, 14)
Sheets(ela).Cells(m + 7, 2) = Sheets("d_P").Cells(3, 10) * 0.65
Sheets(ela).Cells(m + 7, 5) = Sheets("d_P").Cells(3, 14) * 1.35
Sheets(ela).Cells(m + 8, 2) = Sheets("d_P").Cells(3, 10) * 0.8
Sheets(ela).Cells(m + 8, 5) = Sheets("d_P").Cells(3, 14) * 1.2
Sheets(ela).Cells(m + 9, 2) = Sheets("d_P").Cells(3, 10) * 0.65
Sheets(ela).Cells(m + 9, 5) = Sheets("d_P").Cells(3, 14) * 1.35
Sheets(ela).Cells(m + 10, 2) = Sheets("d_P").Cells(3, 10) * 0.8
Sheets(ela).Cells(m + 10, 5) = Sheets("d_P").Cells(3, 14) * 1.2


'读取各时程下基底剪力，汇总至表格
For i = 1 To m

    '基底剪力汇总
    Sheets(ela).Cells(5 + i, 2).Value = Sheets(ela).Cells(3, 11 + (i - 1) * 6)
    Sheets(ela).Cells(5 + i, 5).Value = Sheets(ela).Cells(3, 14 + (i - 1) * 6)

    '时程结果与反应谱结果的比值
    If Sheets(ela).Cells(m + 6, 2) = "" Then
        'Debug.Print "缺少反应谱数据！"
    Else
        Sheets(ela).Cells(5 + i, 3).Value = Round(Sheets(ela).Cells(5 + i, 2).Value / Sheets(ela).Cells(m + 6, 2).Value, 3)
        Sheets(ela).Cells(5 + i, 6).Value = Round(Sheets(ela).Cells(5 + i, 5).Value / Sheets(ela).Cells(m + 6, 5).Value, 3)
    End If
    
    '位移角汇总
    '最大位移角所在行数
    Dim RRX, RRY As Integer
    RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (i - 1) * 6), Sheets(ela).Cells(NN + 2, 10 + (i - 1) * 6)))(1)
    Debug.Print "test" & RRX
    RRY = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 13 + (i - 1) * 6), Sheets(ela).Cells(NN + 2, 13 + (i - 1) * 6)))(1)
    '将最大位移角及构件编号写入表格
    Sheets(ela).Cells(5 + i, 4) = Worksheets(ela).Cells(RRX, 10 + (i - 1) * 6)
    Worksheets(ela).Cells(RRX, 10 + (i - 1) * 6).Interior.ColorIndex = 4
    Sheets(ela).Cells(5 + i, 7) = Worksheets(ela).Cells(RRY, 13 + (i - 1) * 6)
    Worksheets(ela).Cells(RRY, 13 + (i - 1) * 6).Interior.ColorIndex = 4

    '时程结果与反应谱结果的比值
    'Sheets(ela).Cells(6 + i, 3).Value = Round(Sheets(ela).Cells(6 + i, 2).Value / Sheets(ela).Cells(6, 2).Value, 3)
    'Sheets(ela).Cells(6 + i, 6).Value = Round(Sheets(ela).Cells(6 + i, 5).Value / Sheets(ela).Cells(6, 5).Value, 3)
    'Sheets(ela).Cells(20 + i, 3).Value = Round(Sheets(ela).Cells(20, 2).Value / Sheets(ela).Cells(20 + i, 2).Value, 3)
    'Sheets(ela).Cells(20 + i, 6).Value = Round(Sheets(ela).Cells(20, 5).Value / Sheets(ela).Cells(20 + i, 5).Value, 3)
    
Next i

'读取反应谱的最大位移角
If Sheets(ela).Cells(m + 6, 2) = "" Then
MsgBox "缺少反应谱数据！"
Else
RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 10 + m * 6), Sheets(ela).Cells(NN + 2, 10 + m * 6)))(1)
RRY = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 13 + m * 6), Sheets(ela).Cells(NN + 2, 13 + m * 6)))(1)
Sheets(ela).Cells(m + 6, 4) = Worksheets(ela).Cells(RRX, 10 + m * 6)
Worksheets(ela).Cells(RRX, 10 + m * 6).Interior.ColorIndex = 4
Sheets(ela).Cells(m + 6, 7) = Worksheets(ela).Cells(RRY, 13 + m * 6)
Worksheets(ela).Cells(RRY, 13 + m * 6).Interior.ColorIndex = 4
End If

'平均值与反应谱值的比值
'Sheets(ela).Cells(7, 4).Value = Round(Sheets(ela).Cells(5 + m, 2).Value / Sheets(ela).Cells(6, 2).Value, 3)
'Sheets(ela).Cells(7, 7).Value = Round(Sheets(ela).Cells(5 + m, 5).Value / Sheets(ela).Cells(6, 5).Value, 3)

'Sheets(ela).Cells(21, 4).Value = Round(Sheets(ela).Cells(19 + m, 2).Value / Sheets(ela).Cells(20, 2).Value, 3)
'Sheets(ela).Cells(21, 7).Value = Round(Sheets(ela).Cells(19 + m, 5).Value / Sheets(ela).Cells(20, 5).Value, 3)

'=============================================================================================================================格式补充

'Sheets(ela).range("D7:D" & m + 4).MergeCells = True
'Sheets(ela).range("G7:G" & m + 4).MergeCells = True

'加背景色
Call AddShadow(ela, "A2:A" & m + 10, 10092441)
Call AddShadow(ela, "B4:G5,C3:C3,E3:E3", 10092441)
Call AddShadow(ela, "B2:B3,D3:D3,F3:F3", 6750207)
Call AddShadow(ela, "B6:G" & m + 10, 6750207)

'所有单元格宽度自动调整
Sheets(ela).Cells.EntireColumn.AutoFit


'总耗时
Debug.Print "耗费时间: " & Timer - sngStart


'接续最开始的if
'End If

End Sub

Function IndexMinofRange(index_Range As range)
Dim Min, R, C As Integer
Min = WorksheetFunction.Min(index_Range)
R = index_Range.Find(Min, After:=index_Range.Cells(index_Range.Rows.Count, index_Range.Columns.Count), LookIn:=xlValues).Row
C = index_Range.Find(Min, After:=index_Range.Cells(index_Range.Rows.Count, index_Range.Columns.Count), LookIn:=xlValues).Columns
IndexMinofRange = Array(Min, R, C)
End Function


