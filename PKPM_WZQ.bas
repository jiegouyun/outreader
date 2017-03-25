Attribute VB_Name = "PKPM_WZQ"
Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/19/ 9:16
'更新内容:
'1.首层地震作用下的剪力和弯矩写入general
'2.删除对distribution表格格式的设置

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/16/ 17:22
'更新内容:
'1.添加周期比大于或小于0.85

'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/15 17:14
'更新内容：
'1.针对新的general表格更新内容，Z向质量参与系数目前填在计算振型个数位置，计算振型数未读取。
'2.周期比后面两个单元格取消合并，因周期比按数组公式计算，合并单元格不能计算


'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/25 22:14
'更新内容：
'1.楼层总数由文件自身读取


'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/7
'更新内容：
'1.全局变量已经设置，取消这里的设置；


'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/6 23:48
'更新内容：
'1.添加读取平动系数组成
'2.读取最小剪重比时不考虑地下室楼层，地下室楼层数需改成从外传递进来

'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/20 22:38
'更新内容：
'1.读取Z向质量参与系数
'2.判断第一扭转周期，计算周期比
'////////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            PKPM_WZQ.OUT部分代码                      ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_PKPM_WZQ(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring  As String

'定义data为读入行的字符串
Dim data As String


'定义循环变量
Dim i, j As Integer

Dim i_m, i_k1, i_k2, i_w As Integer

'i_k1、i_k2分别为两种刚度比的写入行数记录，第3行为第1层，前两行为标题行
i_k1 = 3
i_k2 = 3

'定义地下室层数,总楼层数
Dim num_floor As Integer
num_floor = 0

'文本当前行数
Dim n As Integer


'==========================================================================================定义关键词变量


'周期
Dim Keyword_Period As String
'赋值
Keyword_Period = "振型号    周 期"

'质量参与系数
Dim Keyword_Mass_Ratio_X, Keyword_Mass_Ratio_Y, Keyword_Mass_Ratio_Z As String
'赋值
Keyword_Mass_Ratio_X = "X 方向的有效质量系数:"
Keyword_Mass_Ratio_Y = "Y 方向的有效质量系数:"
Keyword_Mass_Ratio_Z = "Z 方向的有效质量系数:"


'地震作用
Dim Keyword_Earthquake_X, Keyword_Earthquake_Y As String
'赋值
Keyword_Earthquake_X = "各层 X 方向的作用力(CQC)"
Keyword_Earthquake_Y = "各层 Y 方向的作用力(CQC)"

'最小剪重比
Dim Keyword_Shear_Weight_X, Keyword_Shear_Weight_Y As String
'赋值
Keyword_Shear_Weight_X = "X向楼层最小剪重比"
Keyword_Shear_Weight_Y = "Y向楼层最小剪重比"

'调整后剪重比
Dim Keyword_Shear_Weight_Ratio As String
'赋值
Keyword_Shear_Weight_Ratio = "各楼层地震剪力系数调整情况"


'==========================================================================================定义首字符变量


'周期
Dim Firststring_Period As String

'质量参与系数
Dim Firststring_Mass_Ratio As String

'地震作用
Dim Firststring_Earthquake As String

'最小剪重比
Dim Firststring_Shear_Weight As String

'调整后剪重比
Dim Firststring_Shear_Weight_Ratio As String


'=============================================================================================================================生成文件读取路径

'指定文件名为WZQ.out
Filename = "WZQ.OUT"

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件WZQ.out"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    '记录行数
    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符
   
    '周期
    Firststring_Period = Mid(data, 5, 10)
    '质量参与系数
    Firststring_Mass_Ratio = Mid(data, 8, 12)
    '地震作用
    Firststring_Earthquake = Mid(data, 4, 16)
    '最小剪重比
    Firststring_Shear_Weight = Mid(data, 23, 9)
    '最小剪重比
    Firststring_Shear_Weight_Ratio = Mid(data, 12, 13)

   

    '-------------------------------------------------------------------------------------------读取周期

    If Firststring_Period = Keyword_Period Then

        '按表格格式，从7行开始写入数据
        j = 28

        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){3}") = True Then
                '逐一写入周期、转角、平动系数、扭转系数
                Sheets("g_P").Cells(j, 4) = extractNumberFromString(data, 2)
                Sheets("g_P").Cells(j, 5) = extractNumberFromString(data, 3)
                'Sheets("g_P").Cells(j, 6) = extractNumberFromString(data, 4)
                Sheets("g_P").Cells(j, 6) = StringfromStringforReg(data, "\(.*\)", 1)
                Sheets("g_P").Cells(j, 7) = extractNumberFromString(data, 6)

                '换行
                j = j + 1

            End If

            '遇到分隔符“==”则退出小循环
            If j = 38 Then

                '输出提示语
                Debug.Print "读取周期信息"
                Debug.Print "用时: " & Timer - sngStart
                Debug.Print "……"

                Exit Do

            End If

        Loop

    End If
    
    '-------------------------------------------------------------------------------------------读取质量参与系数


    Select Case Firststring_Mass_Ratio

        Case Keyword_Mass_Ratio_X
        
            '提取质量参与系数数据，并写入工作表相应位置
            Sheets("g_P").Cells(39, 5) = extractNumberFromString(data, 1) '& "%"

            Debug.Print "读取X向质量参与系数"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"

        Case Keyword_Mass_Ratio_Y

            '提取质量参与系数数据，并写入工作表相应位置
            Sheets("g_P").Cells(39, 7) = extractNumberFromString(data, 1) '& "%"

            Debug.Print "读取Y向质量参与系数"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"
            
        'Case Keyword_Mass_Ratio_Z

            '提取质量参与系数数据，并写入工作表相应位置
            'Sheets("g_P").Cells(38, 7) = extractNumberFromString(data, 1)

            'Debug.Print "读取Z向质量参与系数"
            'Debug.Print "用时: " & Timer - sngStart
            'Debug.Print "……"

    End Select
    
    '-------------------------------------------------------------------------------------------读取地震作用


    Select Case Firststring_Earthquake
    
    Case Keyword_Earthquake_X
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "=") Then
            
                '提取最小剪重比规范限值，并写入工作表相应位置
                Sheets("g_P").Cells(24, 7) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '逐一写入层号、剪力X、弯矩X、剪重比X
                Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                Sheets("d_P").Cells(j, 10) = StringfromStringforReg(data, "\d*\.\d*", 2)
                Sheets("d_P").Cells(j, 11) = StringfromStringforReg(data, "\d*\.\d*", 5)
                Sheets("d_P").Cells(j, 12) = StringfromStringforReg(data, "\d*\.\d*", 4)
                '记录总楼层数
                num_floor = num_floor + 1
            
            End If
            
        Loop
        
        Debug.Print "读取X向地震作用"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    Case Keyword_Earthquake_Y
    
        Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "=") Then
                
                    '提取最小剪重比规范限值，并写入工作表相应位置
                    Sheets("g_P").Cells(25, 7) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '逐一写入层号、剪力Y、弯矩Y、剪重比Y
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                    Sheets("d_P").Cells(j, 14) = StringfromStringforReg(data, "\d*\.\d*", 2)
                    Sheets("d_P").Cells(j, 15) = StringfromStringforReg(data, "\d*\.\d*", 5)
                    Sheets("d_P").Cells(j, 16) = StringfromStringforReg(data, "\d*\.\d*", 4)
                
                End If
                
            Loop
            
            Debug.Print "读取Y向地震作用"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"

        End Select
        

    '-------------------------------------------------------------------------------------------读取调整后剪重比

    If Firststring_Shear_Weight_Ratio = Keyword_Shear_Weight_Ratio Then

        Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "本文件结果是在地震外力CQC下的统计结果") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '逐一写入层号、调整后剪重比X、调整后剪重比Y
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                    Sheets("d_P").Cells(j, 13) = StringfromStringforReg(data, "\d*\.\d*", 1) * Sheets("d_P").Cells(j, 12)
                    Sheets("d_P").Cells(j, 17) = StringfromStringforReg(data, "\d*\.\d*", 2) * Sheets("d_P").Cells(j, 16)
                
                End If
                
        Loop
        
        Debug.Print "读取调整后剪重比"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    
        
Loop

'关闭结果文件WZQ.OUT
Close #i
    
    
'-------------------------------------------------------------------------------------------读取最小剪重比
Sheets("g_P").Cells(24, 5).Formula = "=MIN(d_P!L" & CStr(Num_Base + 3) & ":L" & num_floor + 2 & ")"
Sheets("g_P").Cells(25, 5).Formula = "=MIN(d_P!P" & CStr(Num_Base + 3) & ":P" & num_floor + 2 & ")"

'-------------------------------------------------------------------------------------------计算周期比
Sheets("g_P").Cells(38, 4).FormulaArray = "=INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37>0.5,))/INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37<0.5,))"
Sheets("g_P").Cells(38, 5).Formula = "=if(d38<0.85,""< 0.85"",""> 0.85"")"
    
'-------------------------------------------------------------------------------------------读取首层地震作用下的剪力和弯矩
'X向剪力
Sheets("g_P").Cells(44, 4).Formula = "=d_P!J" & Num_Base + 3
'X向弯矩
Sheets("g_P").Cells(44, 6).Formula = "=d_P!K" & Num_Base + 3
'Y向剪力
Sheets("g_P").Cells(45, 4).Formula = "=d_P!N" & Num_Base + 3
'Y向弯矩
Sheets("g_P").Cells(45, 6).Formula = "=d_P!O" & Num_Base + 3

'Sheets("g_P").Cells.EntireColumn.AutoFit
'Sheets("d_P").Cells.EntireColumn.AutoFit
'Sheets("d_P").Cells.NumberFormatLocal = "G/通用格式"

Debug.Print "耗费时间: " & Timer - sngStart

End Sub


