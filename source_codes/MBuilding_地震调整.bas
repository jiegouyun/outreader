Attribute VB_Name = "MBuilding_地震调整"
Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////


'更新时间:2013/7/29 19:19
'更新内容：
'1.增加调整后剪重比的读取

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/20 13:56
'更新内容：
'1.移植PKPM_WV02Q代码，局部作调整

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****           MBuilding_楼层地震作用调整系数.TXT部分代码                 ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_MBuilding_地震调整(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, data  As String

Dim i As Integer, j As Integer

'定义循环变量

Dim i_ As Integer: i = FreeFile


'==========================================================================================定义关键词变量
'柱剪力
Dim Keyword_VC As String
Keyword_VC = "框架承担的地震剪力比"

Dim Keyword_VCX As String
Keyword_VCX = "RS_0"

Dim Keyword_VCY As String
Keyword_VCY = "RS_90"

'剪力调整系数
Dim Keyword_VT As String
Keyword_VT = "0.2Q0调整系数"

Dim Keyword_VTX As String
Keyword_VTX = "RS_0"

Dim Keyword_VTY As String
Keyword_VTY = "RS_90"

'剪重比调整
Dim Keyword_Shear_Weight_Ratio As String
Keyword_Shear_Weight_Ratio = "剪重比调整系数"

'==========================================================================================定义首字符变量
'柱剪力
Dim FirstString_VC As String
Dim FirstString_VCX As String
Dim FirstString_VCY As String
'剪力调整系数
Dim FirstString_VT As String
Dim FirstString_VTX As String
Dim FirstString_VTY As String
'剪重比调整
Dim Firststring_Shear_Weight_Ratio As String

'==========================================================================================生成文件读取路径

'指定文件名
Filename = Dir(Path & "\*_楼层地震作用调整系数.txt")

'生成完整文件路径
filepath = Path & "\" & Filename
'Debug.Print path
'Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件：楼层地震作用调整系数.txt："
Debug.Print "读取相关数据"
Debug.Print "……"

Do While Not EOF(1)

    Line Input #i, data '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = data

    '--------------------------------------------------------------------------定义各指标的判别字符
    FirstString_VC = Mid(data, 6, 10)
    FirstString_VT = Mid(data, 6, 9)
    Firststring_Shear_Weight_Ratio = Mid(data, 6, 7)
   
   
    '--------------------------------------------------------------------------读取柱剪力及其所占总剪力的百分比

    '寻找楼层数
    Dim NA As Integer
        
    If FirstString_VC = Keyword_VC Then
        
        Debug.Print "读取柱剪力及其所占总剪力的百分比……"
        
        '跨过第一个=====行
        Line Input #i, data

        Do While Not EOF(1)
        
            Line Input #i, data
            
            If CheckRegExpfromString(data, "======") Then
                '退出
                Exit Do
            End If
            
            '定义指标判别字符
            FirstString_VCX = Mid(data, 10, 4)
            FirstString_VCY = Mid(data, 10, 5)
            
            '读取X向框架柱分担剪力比
            If FirstString_VCX = Keyword_VCX Then
                
                '跨过第一个-----行
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                
                Do While Not EOF(1)
                
                    Line Input #i, data
                    
                    If CheckRegExpfromString(data, "----") Then
                        '退出
                        Exit Do
                    End If
                    
                    '查询楼层
                    If CheckRegExpfromString(data, "B\S\F") = False Then
                        NA = extractNumberFromString(data, 1) + 2 + Num_Base
                    Else
                        NA = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                    End If
                    'Debug.Print NA
                    '读取框架柱剪力
                    Sheets("d_M").Cells(NA, 48) = StringfromStringforReg(data, "\S+", 4)
                    '求解框架柱剪力承担百分比：（层剪力 - 墙剪力） / 层剪力
                    Sheets("d_M").Cells(NA, 49) = Format((StringfromStringforReg(data, "\S+", 3) - StringfromStringforReg(data, "\S+", 6)) / StringfromStringforReg(data, "\S+", 3) * 100, "0.00")
                    '求解框架柱剪力承担百分比：柱剪力 / 层剪力
                    'Sheets("d_M").Cells(NA + 2, 49) = StringfromStringforReg(data, "\S+", 4) / StringfromStringforReg(data, "\S+", 3)
                Loop
            End If
            
            '读取Y向框架柱分担剪力比
            If FirstString_VCY = Keyword_VCY Then
                
                '跨过第一个-----行
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                
                Do While Not EOF(1)
                
                    Line Input #i, data
                    
                    If CheckRegExpfromString(data, "----") Then
                        '退出
                        Exit Do
                    End If
                    
                    '查询楼层
                    If CheckRegExpfromString(data, "B\S\F") = False Then
                        NA = extractNumberFromString(data, 1) + 2 + Num_Base
                    Else
                        NA = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                    End If
                    'Debug.Print NA
                    '读取框架柱剪力
                    Sheets("d_M").Cells(NA, 51) = StringfromStringforReg(data, "\S+", 4)
                    '求解框架柱剪力承担百分比：（层剪力 - 墙剪力） / 层剪力
                    Sheets("d_M").Cells(NA, 52) = Format((StringfromStringforReg(data, "\S+", 3) - StringfromStringforReg(data, "\S+", 6)) / StringfromStringforReg(data, "\S+", 3) * 100, "0.00")
                    '求解框架柱剪力承担百分比：柱剪力 / 层剪力
                    'Sheets("d_M").Cells(NA + 2, 52) = StringfromStringforReg(data, "\S+", 4) / StringfromStringforReg(data, "\S+", 3)
                Loop
            End If
            
        Loop
        
        Debug.Print "读取柱剪力及其所占总剪力的百分比耗费时间: " & Timer - sngStart
        
    End If

    '--------------------------------------------------------------------------读取X/Y向剪力调整系数
    '寻找层数
    Dim NA2 As Integer
    
    If FirstString_VT = Keyword_VT Then

        Debug.Print "读取柱剪力调整系数……"
        
        '跨过第一个====行
        Line Input #i, data
        
        Do While Not EOF(1)
        
            Line Input #i, data
            
            If CheckRegExpfromString(data, "======") Then
                '退出
                Exit Do
            End If
            
            '定义指标判别字符
            FirstString_VTX = Mid(data, 10, 4)
            FirstString_VTY = Mid(data, 10, 5)
            
            '读取X向剪力调整系数
            If FirstString_VTX = Keyword_VTX Then
                
                '跨过第一个------行
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                
                Do While Not EOF(1)
                
                    Line Input #i, data
                    
                    If CheckRegExpfromString(data, "----") Then
                        '退出
                        Exit Do
                    End If
                    
                    '查询楼层
                    If CheckRegExpfromString(data, "B\S\F") = False Then
                        NA2 = extractNumberFromString(data, 1) + 2 + Num_Base
                    Else
                        NA2 = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                    End If
                    'NA2 = extractNumberFromString(data, 1)
                    'Debug.Print NA2
                    '读取剪力调整系数
                    Sheets("d_M").Cells(NA2, 50) = StringfromStringforReg(data, "\S+", 5)
                                   
                Loop
            End If
            
            '读取X向剪力调整系数
            If FirstString_VTY = Keyword_VTY Then
                
                '跨过第一个------行
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                Line Input #i, data
                
                Do While Not EOF(1)
                
                    Line Input #i, data
                    
                    If CheckRegExpfromString(data, "----") Then
                        '退出
                        Exit Do
                    End If
                    If CheckRegExpfromString(data, "B\S\F") = False Then
                        NA2 = extractNumberFromString(data, 1) + 2 + Num_Base
                    Else
                        NA2 = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                    End If
                     '查询楼层
                    'NA2 = extractNumberFromString(data, 1)
                    'Debug.Print NA2
                    '读取剪力调整系数
                    Sheets("d_M").Cells(NA2, 53) = StringfromStringforReg(data, "\S+", 5)
                                    
                Loop
            End If
                        
        Loop
        
        Debug.Print "读取柱剪力调整系数耗费时间: " & Timer - sngStart
        
    End If
    
    
    '-------------------------------------------------------------------------------------------读取调整后剪重比

    If Firststring_Shear_Weight_Ratio = Keyword_Shear_Weight_Ratio Then
        Debug.Print "读取调整后剪重比……"
        
        Do While Not EOF(1)
            Line Input #i, data
                
                If Mid(data, 3, 4) = "工况 1" Then
                
                    Do While Not EOF(1)
                    Line Input #i, data
                    
                    If Mid(data, 3, 4) = "工况 2" Then
                        Exit Do
                    End If
                    
                '如果接连两个数，认为是数据行
                    If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                        '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                        'j为读取行数据写入表格的行数，跳过两行标题行
                        If CheckRegExpfromString(data, "B\S\F") = False Then
                            j = extractNumberFromString(data, 1) + 2 + Num_Base
                            Sheets("d_M").Cells(j, 13) = StringfromStringforReg(data, "\d*\.\d*", 4) * Sheets("d_M").Cells(j, 12)
                        Else
                            j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                            Sheets("d_M").Cells(j, 13) = Sheets("d_M").Cells(j, 12) * extractNumberFromString(data, 3)
                        End If
                        '逐一写入调整后剪重比X
                        'Sheets("d_M").Cells(j, 13) = StringfromStringforReg(data, "\d*\.\d*", 4) * Sheets("d_M").Cells(j, 12)
                
                    End If
                    Loop
                End If
                
                '如果接连两个数，认为是数据行
                    If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){4}") Then
                
                        '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                        'j为读取行数据写入表格的行数，跳过两行标题行
                        If CheckRegExpfromString(data, "B\S\F") = False Then
                            j = extractNumberFromString(data, 1) + 2 + Num_Base
                            Sheets("d_M").Cells(j, 17) = StringfromStringforReg(data, "\d*\.\d*", 4) * Sheets("d_M").Cells(j, 16)
                        Else
                            j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                            Sheets("d_M").Cells(j, 17) = Sheets("d_M").Cells(j, 16) * extractNumberFromString(data, 3)
                        End If
                    
                        '逐一写入调整后剪重比Y
                        'Sheets("d_M").Cells(j, 17) = StringfromStringforReg(data, "\d*\.\d*", 4) * Sheets("d_M").Cells(j, 16)
                
                    End If
                
        Loop
        
        Debug.Print "读取调整后剪重比"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If


Loop

Close #i


Debug.Print "读取柱剪力信息耗费时间: " & Timer - sngStart


End Sub


