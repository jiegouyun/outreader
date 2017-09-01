Attribute VB_Name = "MBuilding_周期振型"
Option Explicit


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/31 14:00
'更新内容:
'1.修改平动系数的格式
'2.补充振型参与质量系数的写入

'/////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/29 14:19
'更新内容:
'1.移植YJK的prebeta版

'////////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                  MBuilding_周期、地震作用及振型部分代码              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_MBuilding_周期振型(Path As String)

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
'X、Y方向的平动系数
Dim Cof_Px, Cof_Py As Single


'==========================================================================================定义关键词变量


'周期
Dim Keyword_Period As String
'赋值
Keyword_Period = "振型号       周  期"

'质量参与系数
Dim Keyword_Mass_Ratio_X, Keyword_Mass_Ratio_Y, Keyword_Mass_Ratio_Z As String
'赋值
Keyword_Mass_Ratio_X = "X向平动振型参与质量系数总计"
Keyword_Mass_Ratio_Y = "Y向平动振型参与质量系数总计"
Keyword_Mass_Ratio_Z = "Z向扭转振型参与质量系数总计"


'地震作用
Dim Keyword_Earthquake_X, Keyword_Earthquake_Y As String
'赋值
Keyword_Earthquake_X = "[RS_0] 各层地震作用 (CQC(耦联))"
Keyword_Earthquake_Y = "[RS_90] 各层地震作用 (CQC(耦联))"

'最小剪重比
Dim Keyword_Shear_Weight_X, Keyword_Shear_Weight_Y As String
'赋值
Keyword_Shear_Weight_X = "抗震规范(5.2.5条)中要求的最小剪重比"
Keyword_Shear_Weight_Y = "抗震规范(5.2.5条)中要求的最小剪重比"

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
Dim Firststring_Earthquake_X As String, Firststring_Earthquake_Y As String

'最小剪重比
Dim Firststring_Shear_Weight As String

'调整后剪重比
Dim Firststring_Shear_Weight_Ratio As String


'=============================================================================================================================生成文件读取路径

'指定文件名为WDISP.out
Filename = Dir(Path & "\*_周期、地震作用及振型.txt")

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果" & Filename
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
    Firststring_Period = Mid(data, 3, 14)
    '质量参与系数
    Firststring_Mass_Ratio = Mid(data, 5, 14)
    '地震作用
    Firststring_Earthquake_X = Mid(data, 3, 23)
    Firststring_Earthquake_Y = Mid(data, 3, 24)
    '最小剪重比
    Firststring_Shear_Weight = Mid(data, 3, 21)
    '调整后剪重比
    'Firststring_Shear_Weight_Ratio = Mid(data, 12, 13)

   

    '-------------------------------------------------------------------------------------------读取周期

    If Firststring_Period = Keyword_Period Then

        '按表格格式，从7行开始写入数据
        j = 28

        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){3}") = True Then
                '逐一写入周期、转角、平动系数、扭转系数
                Sheets("g_M").Cells(j, 4) = extractNumberFromString(data, 2)
                Sheets("g_M").Cells(j, 5) = "-"
                'Sheets("g_M").Cells(j, 6) = extractNumberFromString(data, 4)
                Cof_Px = Round(extractNumberFromString(data, 3) / 100, 2)
                Cof_Py = Round(extractNumberFromString(data, 4) / 100, 2)
                Sheets("g_M").Cells(j, 6) = "(" & Cof_Px & "+" & Cof_Py & ")"
                Sheets("g_M").Cells(j, 7) = StringfromStringforReg(data, "\S+", 5) / 100

                '换行
                j = j + 1

            End If

            '读够10行，则退出
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
            Sheets("g_M").Cells(39, 5) = extractNumberFromString(data, 1)

            Debug.Print "读取X向质量参与系数"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"

        Case Keyword_Mass_Ratio_Y

            '提取质量参与系数数据，并写入工作表相应位置
            Sheets("g_M").Cells(39, 7) = extractNumberFromString(data, 1)

            Debug.Print "读取Y向质量参与系数"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"
            
       ' Case Keyword_Mass_Ratio_Z

            '提取质量参与系数数据，并写入工作表相应位置
            'Sheets("g_M").Cells(38, 7) = extractNumberFromString(data, 1)

            'Debug.Print "读取Z向质量参与系数"
            'Debug.Print "用时: " & Timer - sngStart
            'Debug.Print "……"

    End Select
    
    '-------------------------------------------------------------------------------------------读取地震作用


    Select Case Firststring_Earthquake_X
    
    Case Keyword_Earthquake_X
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "=") Then
            
                '提取最小剪重比规范限值，并写入工作表相应位置
                Sheets("g_M").Cells(24, 7) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                
                '逐一写入剪力X、弯矩X、剪重比X
'                Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                Sheets("d_M").Cells(j, 10) = StringfromStringforReg(data, "\d*\.\d*", 2)
                Sheets("d_M").Cells(j, 11) = StringfromStringforReg(data, "\d*\.\d*", 4)
                Sheets("d_M").Cells(j, 12) = StringfromStringforReg(data, "\d*\.\d*", 3)
                '记录总楼层数
                num_floor = num_floor + 1
            
            End If
            
        Loop
        
        Debug.Print "读取X向地震作用"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"
        End Select
        
    Select Case Firststring_Earthquake_Y
    Case Keyword_Earthquake_Y
    
        Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "=") Then
                
                    '提取最小剪重比规范限值，并写入工作表相应位置
                    Sheets("g_M").Cells(25, 7) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                    
                    '逐一写入剪力Y、弯矩Y、剪重比Y
                    'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                    Sheets("d_M").Cells(j, 14) = StringfromStringforReg(data, "\d*\.\d*", 2)
                    Sheets("d_M").Cells(j, 15) = StringfromStringforReg(data, "\d*\.\d*", 4)
                    Sheets("d_M").Cells(j, 16) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
            
            Debug.Print "读取Y向地震作用"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"

        End Select
        
    
        
Loop

'关闭结果文件
Close #i
    
    
'-------------------------------------------------------------------------------------------读取最小剪重比
Sheets("g_M").Cells(24, 5).Formula = "=MIN(d_M!L" & CStr(Num_Base + 3) & ":L" & num_floor + 2 & ")"
Sheets("g_M").Cells(25, 5).Formula = "=MIN(d_M!P" & CStr(Num_Base + 3) & ":P" & num_floor + 2 & ")"

'-------------------------------------------------------------------------------------------计算周期比
Sheets("g_M").Cells(38, 4).FormulaArray = "=INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37>0.5,))/INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37<0.5,))"
Sheets("g_M").Cells(38, 5).Formula = "=if(d38<0.85,""< 0.85"",""> 0.85"")"

'-------------------------------------------------------------------------------------------读取首层地震作用下的剪力和弯矩
'X向剪力
Sheets("g_M").Cells(44, 4).Formula = "=d_M!J" & Num_Base + 3
'X向弯矩
Sheets("g_M").Cells(44, 6).Formula = "=d_M!K" & Num_Base + 3
'Y向剪力
Sheets("g_M").Cells(45, 4).Formula = "=d_M!N" & Num_Base + 3
'Y向弯矩
Sheets("g_M").Cells(45, 6).Formula = "=d_M!O" & Num_Base + 3


'Sheets("g_M").Cells.EntireColumn.AutoFit
'Sheets("d_M").Cells.EntireColumn.AutoFit
'Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"

Debug.Print "耗费时间: " & Timer - sngStart

End Sub

