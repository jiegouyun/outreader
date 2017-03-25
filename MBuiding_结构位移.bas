Attribute VB_Name = "MBuiding_结构位移"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/11
'更新内容:
'1.添加range是否为空判断,解决缺少数据时with报错问题


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'更新内容:
'1.添加数据格式代码,解决with报错问题

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/04/18
'更新内容:
'1.位移比取最值前添加format限制代码

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15

'更新内容：
'1.修正MIDAS考虑扭转时风荷载工况下位移角的提取；
'2.修改高亮代码；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29

'更新内容：
'1.增加高亮代码

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12 17:13

'更新内容：
'1.修改最大层间位移角取值范围，只限于地震与风，不含偏心
'2.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/31 13:58
'更新内容:
'1.更改位移比区域数据格式为小数点后三位，修正位移比取大值时出错

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/29 12:48
'更新内容:
'1.移植PKPM的prebeta版

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/12 14:42
'更新内容:
'1.MBuilding层间位移角分母最大可能大于9999，考虑读取分母为5位的情况



'///////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            MBuilding_结构位移部分代码                ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_MBuilding_结构位移(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring  As String


'定义data为读入行的字符串
Dim data As String


'定义循环变量
Dim i As Integer, j As Integer

Dim i_m As Integer, i_k1 As Integer, i_k2 As Integer, i_w As Integer, i_q As Integer

'i_k1、i_k2分别为两种刚度比的写入行数记录，第3行为第1层，前两行为标题行
i_k1 = 3
i_k2 = 3


'文本当前行数
Dim n As Integer

'定义所在楼层查询区域
Dim iRng_X As range, iRng_Y As range


'==========================================================================================定义关键词变量


'层间位移角
Dim Keyword_Story_Drift_EX As String

Dim Keyword_Story_Drift_EXP As String

Dim Keyword_Story_Drift_EXN As String

Dim Keyword_Story_Drift_EY As String

Dim Keyword_Story_Drift_EYP As String

Dim Keyword_Story_Drift_EYN As String

Dim Keyword_Story_Drift_WX As String

Dim Keyword_Story_Drift_WY As String

'赋值
Keyword_Story_Drift_EX = "RS_0作用下楼层位移"

'Keyword_Story_Drift_EXP = "X+ 偶然偏心地震作用下的楼层最大位移"

'Keyword_Story_Drift_EXN = "X- 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_EY = "RS_90作用下楼层位移"

'Keyword_Story_Drift_EYP = "Y+ 偶然偏心地震作用下的楼层最大位移"

'Keyword_Story_Drift_EYN = "Y- 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_WX = "WL_0作用下"

Keyword_Story_Drift_WY = "WL_90作用"



'位移比
Dim Keyword_Disp_Ratio_FEX As String

Dim Keyword_Disp_Ratio_FEXP As String

Dim Keyword_Disp_Ratio_FEXN As String

Dim Keyword_Disp_Ratio_FEY As String

Dim Keyword_Disp_Ratio_FEYP As String

Dim Keyword_Disp_Ratio_FEYN As String

'赋值
'Keyword_Disp_Ratio_FEX = "X 方向规定水平力作用下的楼层最大位移"

Keyword_Disp_Ratio_FEXP = "RS_0+ES_0作用"

Keyword_Disp_Ratio_FEXN = "RS_0-ES_0作用"

'Keyword_Disp_Ratio_FEY = "Y 方向规定水平力作用下的楼层最大位移"

Keyword_Disp_Ratio_FEYP = "RS_90+ES_90作用"

Keyword_Disp_Ratio_FEYN = "RS_90-ES_90作用"

'==========================================================================================定义首字符变量


'层间位移角/位移比

Dim Firststring_Disp, Firststring_Story_Drift_EX, Firststring_Story_Drift_EXP, Firststring_Story_Drift_EXN, Firststring_Story_Drift_EY, Firststring_Story_Drift_EYP, Firststring_Story_Drift_EYN As String

Dim Firststring_Story_Drift_WX, Firststring_Story_Drift_WY, Firststring_Disp_Ratio_FEX, Firststring_Disp_Ratio_FEXP, Firststring_Disp_Ratio_FEXN, Firststring_Disp_Ratio_FEY, Firststring_Disp_Ratio_FEYP, Firststring_Disp_Ratio_FEYN As String

'=============================================================================================================================生成文件读取路径

'指定文件名为WDISP.out
Filename = Dir(Path & "\*_结构位移.txt")

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件" & Filename
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    '记录行数
    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符
   
    '层间位移角/位移比
    Firststring_Disp = Mid(data, 17, 25)
    
    Firststring_Story_Drift_EX = Mid(data, 6, 11)

    'Firststring_Story_Drift_EXP = Mid(data, 15, 19)

    'Firststring_Story_Drift_EXN = Mid(data, 15, 19)

    Firststring_Story_Drift_EY = Mid(data, 6, 12)

    'Firststring_Story_Drift_EYP = Mid(data, 16, 25)

    'Firststring_Story_Drift_EYN = Mid(data, 16, 25)

    Firststring_Story_Drift_WX = Mid(data, 6, 7)

    Firststring_Story_Drift_WY = Mid(data, 6, 7)

    'Firststring_Disp_Ratio_FEX = Mid(data, 15, 25)

    Firststring_Disp_Ratio_FEXP = Mid(data, 6, 11)

    Firststring_Disp_Ratio_FEXN = Mid(data, 6, 11)

    'Firststring_Disp_Ratio_FEY = Mid(data, 16, 25)

    Firststring_Disp_Ratio_FEYP = Mid(data, 6, 13)

    Firststring_Disp_Ratio_FEYN = Mid(data, 6, 13)
    
    '-------------------------------------------------------------------------------------------读取层间位移角/位移比

    
    '----------------------------------------------------------读取X 方向地震作用下的位移、层间位移角
    If Firststring_Story_Drift_EX = Keyword_Story_Drift_EX Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1)
                
                '写入最大位移
                Sheets("d_M").Cells(j, 18) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 26) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
        
    End If


        '----------------------------------------------------------读取Y 方向地震作用下的位移、层间位移角
    If Firststring_Story_Drift_EY = Keyword_Story_Drift_EY Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1)
                
                '写入最大位移
                Sheets("d_M").Cells(j, 22) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 30) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
    End If
        
    ''----------------------------------------------------------读取X+ 偶然偏心地震作用下的位移、层间位移角
    'If Firststring_Story_Drift_EXP = Keyword_Story_Drift_EXP Then
    '    Debug.Print data
    '    Do While Not EOF(1)
    '        Line Input #i, data
            
    '        If CheckRegExpfromString(data, "层间位移角") Then
    '            '退出
    '            Exit Do
    '        End If
            
    '        '如果接连两个数，认为是数据行
    '        If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
    '            '结果文件中数据对应层号从大至小，统一为从小到大排列
                
    '            'j为读取行数据写入表格的行数，跳过两行标题行
    '            j = extractNumberFromString(data, 1) + 2
                
    '            '读取楼层
    '            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
    '            '写入平均位移
    '            Sheets("d_M").Cells(j, 19) = StringfromStringforReg(data, "\d*\.\d*", 4)
                
                
    '            '读取下一行中层间位移角数据
    '            Line Input #i, data
    '            '写入层间位移角
    '            Sheets("d_M").Cells(j, 27) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
    '        End If
    '    Loop
    ' End If
        
    ''----------------------------------------------------------读取X- 偶然偏心地震作用下的位移、层间位移角
    'If Firststring_Story_Drift_EXN = Keyword_Story_Drift_EXN Then
    '    Debug.Print data
    '    Do While Not EOF(1)
    '        Line Input #i, data
            
    '        If CheckRegExpfromString(data, "最大层间位移角") Then
    '            '退出
    '            Exit Do
    '        End If
            
    '        '如果接连两个数，认为是数据行
    '        If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
    '            '结果文件中数据对应层号从大至小，统一为从小到大排列
                
    '            'j为读取行数据写入表格的行数，跳过两行标题行
    '            j = extractNumberFromString(data, 1) + 2
                
    '            '读取楼层
    '            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
    '            '写入平均位移
    '            Sheets("d_M").Cells(j, 20) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
    '            '读取下一行中层间位移角数据
    '            Line Input #i, data
    '            '写入层间位移角
    '            Sheets("d_M").Cells(j, 28) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
    '        End If
    '    Loop
    ' End If
        
    ''----------------------------------------------------------读取Y+ 偶然偏心地震作用下的位移、层间位移角
    'If Firststring_Story_Drift_EYP = Keyword_Story_Drift_EYP Then
    '    Debug.Print data
    '    Do While Not EOF(1)
    '        Line Input #i, data
            
    '        If CheckRegExpfromString(data, "最大层间位移角") Then
    '            '退出
    '            Exit Do
    '        End If
            
    '        '如果接连两个数，认为是数据行
    '        If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
    '            '结果文件中数据对应层号从大至小，统一为从小到大排列
                
    '            'j为读取行数据写入表格的行数，跳过两行标题行
    '            j = extractNumberFromString(data, 1) + 2
                
    '            '读取楼层
    '            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
    '            '写入平均位移
    '            Sheets("d_M").Cells(j, 23) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
    '            '读取下一行中层间位移角数据
    '            Line Input #i, data
    '            '写入层间位移角
    '            Sheets("d_M").Cells(j, 31) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
    '        End If
    '    Loop
        
    ' End If
             
    ''----------------------------------------------------------读取Y- 偶然偏心地震作用下的层间位移角
    'If Firststring_Story_Drift_EYN = Keyword_Story_Drift_EYN Then
    '    Debug.Print data
    '    Do While Not EOF(1)
    '        Line Input #i, data
            
    '        If CheckRegExpfromString(data, "最大层间位移角") Then
    '            '退出
    '            Exit Do
    '        End If
            
    '        '如果接连两个数，认为是数据行
    '        If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
    '            '结果文件中数据对应层号从大至小，统一为从小到大排列
                
    '            'j为读取行数据写入表格的行数，跳过两行标题行
    '            j = extractNumberFromString(data, 1) + 2
                
    '            '读取楼层
    '            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
    '            '写入平均位移
    '            Sheets("d_M").Cells(j, 24) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
    '            '读取下一行中层间位移角数据
    '            Line Input #i, data
    '            '写入层间位移角
    '            Sheets("d_M").Cells(j, 32) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
    '        End If
    '    Loop
        
    ' End If
        
    '----------------------------------------------------------X 方向风荷载作用下的层间位移角
    If Mid(data, 6, 7) = "WL_0作用下" Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
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
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1)
                
                '写入最大位移
                Sheets("d_M").Cells(j, 21) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 29) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
        
     End If
     
    '----------------------------------------------------------X 方向风荷载作用下的层间位移角
    If Mid(data, 6, 7) = "WL_0_DK" Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
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
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1) + Num_Base
                
                '写入最大位移
                Sheets("d_M").Cells(j, 21) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 29) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
        
     End If
        
    '----------------------------------------------------------Y 方向风荷载作用下的层间位移角
    If Firststring_Story_Drift_WY = Keyword_Story_Drift_WY Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
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
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1) + Num_Base
                
                
                '写入最大位移
                Sheets("d_M").Cells(j, 25) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 33) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
        
    End If
    
    '----------------------------------------------------------Y 方向风荷载作用下的层间位移角
    If Firststring_Story_Drift_WY = "WL_90_D" Then
        Debug.Print data
        Do While Not EOF(1)
            Line Input #i, data
            
            If Mid(data, 3, 7) = "最大层间位移角" Then
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
                
                '读取楼层
                'Sheets("d_M").Cells(j, 1) = extractNumberFromString(data, 1)
                
                '写入最大位移
                Sheets("d_M").Cells(j, 25) = StringfromStringforReg(data, "\b\d*\.\d*", 1)
                
                '读取层间位移角数据
                '写入层间位移角
                Sheets("d_M").Cells(j, 33) = Mid(StringfromStringforReg(data, "1/\s*\d+\s", 1), 3, 5)
            
            End If
        Loop
        
    End If
        
        
    ''----------------------------------------------------------读取X 方向地震作用规定水平力下的位移比、层间位移比
    'If Firststring_Disp_Ratio_FEX = Keyword_Disp_Ratio_FEX Then
    '        Debug.Print data
    '        Do While Not EOF(1)
    '            Line Input #i, data
                
    '            If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Or CheckRegExpfromString(data, "最大层间位移角") Then
    '                '退出
    '                Exit Do
    '            End If
                
    '            '如果接连两个数，认为是数据行
    '            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
    '                '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
    '                'j为读取行数据写入表格的行数，跳过两行标题行
    '                j = extractNumberFromString(data, 1) + 2
                    
    '                '读取楼层
    '                'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
    '                '写入位移比
    '                Sheets("d_M").Cells(j, 34) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
    '                '读取下一行中层间位移比数据
    '                Line Input #i, data
    '                '写入层间位移比
    '                Sheets("d_M").Cells(j, 40) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
    '            End If
    '        Loop
            
    '  End If
            
    '----------------------------------------------------------读取X+偶然偏心地震作用规定水平力下的位移比、层间位移比
    If Firststring_Disp_Ratio_FEXP = Keyword_Disp_Ratio_FEXP Then
            Line Input #i, data
            
            Do While Not EOF(1)
                Line Input #i, data
                
                If Mid(data, 3, 11) = "楼层扭转位移比(高规)" Then
                    Do While Not EOF(1)
                        Line Input #i, data
                        
                        '如果接连两个数，认为是数据行
                        If CheckRegExpfromString(data, "Base") Then
                
                            '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                            'j为读取行数据写入表格的行数，跳过两行标题行
                            If CheckRegExpfromString(data, "B\S\F") = False Then
                                j = extractNumberFromString(data, 1) + 2 + Num_Base
                            Else
                                j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                            End If
                    
                            '读取楼层
                            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                            '写入位移比
                            Sheets("d_M").Cells(j, 35) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    

                            '写入层间位移比
                            Sheets("d_M").Cells(j, 41) = StringfromStringforReg(data, "\d*\.\d*", 6)
                
                        End If
                
                
                        If CheckRegExpfromString(data, "==") Then
                            '退出
                            Exit Do
                        End If
                    Loop
                    
                End If
                
                If CheckRegExpfromString(data, "==") Then
                    Exit Do
                End If
            Loop
            
      End If
        
      '----------------------------------------------------------读取X-偶然偏心地震作用规定水平力下的位移比、层间位移比
      If Firststring_Disp_Ratio_FEXN = Keyword_Disp_Ratio_FEXN Then
            Line Input #i, data
            Debug.Print data
            Do While Not EOF(1)
                Line Input #i, data
                
                If Mid(data, 3, 11) = "楼层扭转位移比(高规)" Then
                    Do While Not EOF(1)
                        Line Input #i, data
                
                        '如果接连两个数，认为是数据行
                        If CheckRegExpfromString(data, "Base") Then
                
                             '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                             'j为读取行数据写入表格的行数，跳过两行标题行
                            If CheckRegExpfromString(data, "B\S\F") = False Then
                                j = extractNumberFromString(data, 1) + 2 + Num_Base
                            Else
                                j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                            End If
                    
                            '读取楼层
                            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                            '写入位移比
                            Sheets("d_M").Cells(j, 36) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                            '写入层间位移比
                            Sheets("d_M").Cells(j, 42) = StringfromStringforReg(data, "\d*\.\d*", 6)
                
                        End If
                        
                        If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Then
                            '退出
                            Exit Do
                        End If
                    Loop
                End If
                
                If CheckRegExpfromString(data, "==") Then
                    Exit Do
                End If
            Loop
            
       End If
            
        '----------------------------------------------------------读取Y 方向地震作用规定水平力下的位移比、层间位移比
      'If Firststring_Disp_Ratio_FEY = Keyword_Disp_Ratio_FEY Then
      '      Debug.Print data
      '      Do While Not EOF(1)
      '          Line Input #i, data
                
      '          If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Or CheckRegExpfromString(data, "最大层间位移角") Then
      '              '退出
      '              Exit Do
      '          End If
                
      '          '如果接连两个数，认为是数据行
      '          If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
      '              '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
      '              'j为读取行数据写入表格的行数，跳过两行标题行
      '              j = extractNumberFromString(data, 1) + 2
                    
      '              '读取楼层
      '              'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
      '              '写入位移比
      '              Sheets("d_M").Cells(j, 37) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
      '              '读取下一行中层间位移比数据
      '              Line Input #i, data
      '              '写入层间位移比
      '              Sheets("d_M").Cells(j, 43) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
      '          End If
      '      Loop
            
      '  End If
            
        '----------------------------------------------------------读取Y+偶然偏心地震作用规定水平力下的位移比、层间位移比
        If Firststring_Disp_Ratio_FEYP = Keyword_Disp_Ratio_FEYP Then
            Line Input #i, data
            Debug.Print data
            Do While Not EOF(1)
                Line Input #i, data
                
                If Mid(data, 3, 11) = "楼层扭转位移比(高规)" Then
                    Do While Not EOF(1)
                        Line Input #i, data
                
                        '如果接连两个数，认为是数据行
                        If CheckRegExpfromString(data, "Base") Then
                
                             '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                            'j为读取行数据写入表格的行数，跳过两行标题行
                        If CheckRegExpfromString(data, "B\S\F") = False Then
                            j = extractNumberFromString(data, 1) + 2 + Num_Base
                        Else
                            j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                        End If
                    
                            '读取楼层
                            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                            '写入位移比
                            Sheets("d_M").Cells(j, 38) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                            '写入层间位移比
                            Sheets("d_M").Cells(j, 44) = StringfromStringforReg(data, "\d*\.\d*", 6)
                
                        End If
                        If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Then
                            '退出
                            Exit Do
                        End If
                    Loop
                End If
                
                If CheckRegExpfromString(data, "==") Then
                    Exit Do
                End If

            Loop
            
         End If
        
        '----------------------------------------------------------读取Y-偶然偏心地震作用规定水平力下的位移比、层间位移比
        If Firststring_Disp_Ratio_FEYN = Keyword_Disp_Ratio_FEYN Then
            Line Input #i, data
            Debug.Print data
            Do While Not EOF(1)
                Line Input #i, data
                
                If Mid(data, 3, 11) = "楼层扭转位移比(高规)" Then
                    Do While Not EOF(1)
                        Line Input #i, data
                        
                        '如果接连两个数，认为是数据行
                        If CheckRegExpfromString(data, "Base") Then
                
                            '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                            'j为读取行数据写入表格的行数，跳过两行标题行
                            If CheckRegExpfromString(data, "B\S\F") = False Then
                                j = extractNumberFromString(data, 1) + 2 + Num_Base
                            Else
                                j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                            End If
                    
                            '读取楼层
                            'Sheets("d_M").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                            '写入位移比
                            Sheets("d_M").Cells(j, 39) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                            '写入层间位移比
                            Sheets("d_M").Cells(j, 45) = StringfromStringforReg(data, "\d*\.\d*", 6)
                
                        End If
                        If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Then
                            '退出
                            Exit Do
                        End If
                    Loop
                End If
                
                If CheckRegExpfromString(data, "==") Then
                    Exit Do
                End If

            Loop
        End If

Loop

    Debug.Print "读取地震作用"
    Debug.Print "用时: " & Timer - sngStart
    Debug.Print "……"

'关闭结果文件WDISP.OUT
Close #i

''-------------------------------------------------------------------------------------------高亮最值
'
'Sheets("d_M").Cells.EntireColumn.AutoFit
'
'Num_All = Sheets("d_M").range("a65536").End(xlUp)
'
'
'Dim ii As Integer
'Dim i_RowID As Integer
'Dim i_Rng As range
'
'
''---------------------------------------------------------位移角
'For ii = 26 To 26
'Dim R As range
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "min", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
'For ii = 29 To 30
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "min", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
'For ii = 33 To 33
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "min", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
'
''---------------------------------------------------------位移比
'For ii = 35 To 36
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'For ii = 38 To 39
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'For ii = 41 To 42
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'For ii = 44 To 45
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
'Sheets("g_M").Cells.EntireColumn.AutoFit
'Sheets("d_M").Cells.EntireColumn.AutoFit
'Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"


'-------------------------------------------------------------------------------------------读取最大层间位移角、所在楼层及工况
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  风荷载
Sheets("g_M").Cells(10, 5).Formula = "=1&"" / ""&MIN(d_M!AC:AC)"
Sheets("g_M").Cells(10, 7).Formula = "=1&"" / ""&MIN(d_M!AG:AG)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震
Sheets("g_M").Cells(11, 5).Formula = "=1&"" / ""&MIN(d_M!Z:Z)"
Sheets("g_M").Cells(11, 7).Formula = "=1&"" / ""&MIN(d_M!AD:AD)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震+
'Sheets("g_M").Cells(12, 5).Formula = "=1&"" / ""&MIN(d_M!AA:AA)"
Sheets("g_M").Cells(12, 5).Formula = "-"
'Sheets("g_M").Cells(12, 7).Formula = "=1&"" / ""&MIN(d_M!AE:AE)"
Sheets("g_M").Cells(12, 7).Formula = "-"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震-
'Sheets("g_M").Cells(13, 5).Formula = "=1&"" / ""&MIN(d_M!AB:AB)"
Sheets("g_M").Cells(13, 5).Formula = "-"
'Sheets("g_M").Cells(13, 7).Formula = "=1&"" / ""&MIN(d_M!AF:AF)"
Sheets("g_M").Cells(13, 7).Formula = "-"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  最大层间位移角
Sheets("g_M").Cells(14, 4).Formula = "=1&"" / ""&MIN(d_M!Z:Z,d_M!AC:AD,d_M!AG:AG)"


'//////////////////////////////////////////////////////////////////////////////////////////////
'调用EXCEl公式，计算慢，每次需先清除原先的公式才行
'Sheets("g_M").Cells(18, 14).FormulaArray = "=INDEX(d_M!C[-13],SMALL(IF(d_M!C[12]:C[14]=MIN(d_M!C[12]:C[14]),ROW(R[-17]:R[4982]),4^8),ROW(R[-17])))&"""""
'Sheets("g_M").Cells(18, 15).FormulaArray = "=INDEX(d_M!C[-14],SMALL(IF(d_M!C[14]:C[16]=MIN(d_M!C[14]:C[16]),ROW(R[-17]:R[4982]),4^8),ROW(R[-17])))&"""""
'//////////////////////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////////////////////
'使用VBA查询功能，快速


'定义最大层间位移角查询区域
Set iRng_X = Application.Union(range("d_M!Z:Z"), range("d_M!AC:AD"), range("d_M!AG:AG"))

'定义查询变量
Dim i_Min As Double, i_Row As Integer, i_Col As Integer
Dim i_Temp As range

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大层间位移角（实际是查询层间位移角的最小分母）
i_Min = WorksheetFunction.Min(range("d_M!Z:AG"))

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大层间位移角所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大层间位移角所在层，及其工况
    Sheets("g_M").Cells(15, 7) = Sheets("d_M").Cells(i_Row, 1)
    Sheets("g_M").Cells(15, 5) = Sheets("d_M").Cells(2, i_Col)
End If

'返回最大层间位移角所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大层间位移角所在层，及其工况
'Sheets("g_M").Cells(15, 7) = Sheets("d_M").Cells(i_Row, 1)
'Sheets("g_M").Cells(15, 5) = Sheets("d_M").Cells(2, i_Col)

Sheets("d_M").Columns("AH:AS").NumberFormatLocal = "0.000"
    
'-------------------------------------------------------------------------------------------读取最大位移比
Sheets("g_M").Cells(16, 4).Formula = "=MAX(d_M!AH:AM)"

'------------------------------------------------------------------------------------------- 查询最大位移比所在楼层

'定义最大位移比查询区域
Set iRng_X = range("d_M!AH:AM")
'MBuilding位移比原始数据小数点后保留三位，更改单元格格式适应，否则查询区域时会出错
Sheets("d_M").range("AH:AS").NumberFormatLocal = "0.000"


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大位移比
i_Min = WorksheetFunction.Max(range("d_M!AH:AM"))
i_Min = Format(i_Min, "0.000")

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大位移比所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大位移比所在层，及其工况
    Sheets("g_M").Cells(17, 7) = Sheets("d_M").Cells(i_Row, 1)
    Sheets("g_M").Cells(17, 5) = Sheets("d_M").Cells(2, i_Col)
End If

'返回最大位移比所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大位移比所在层，及其工况
'Sheets("g_M").Cells(17, 7) = Sheets("d_M").Cells(i_Row, 1)
'Sheets("g_M").Cells(17, 5) = Sheets("d_M").Cells(2, i_Col)


'-------------------------------------------------------------------------------------------读取最大层间位移比
Sheets("g_M").Cells(18, 4).Formula = "=MAX(d_M!AN:AS)"

'------------------------------------------------------------------------------------------- 查询最大层间位移比所在楼层

'定义最大位移比查询区域
Set iRng_X = range("d_M!AN:AS")

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大位移比
i_Min = WorksheetFunction.Max(range("d_M!AN:AS"))
i_Min = Format(i_Min, "0.000")

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大位移比所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大位移比所在层，及其工况
    Sheets("g_M").Cells(19, 7) = Sheets("d_M").Cells(i_Row, 1)
    Sheets("g_M").Cells(19, 5) = Sheets("d_M").Cells(2, i_Col)
End If
    
'返回最大位移比所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大位移比所在层，及其工况
'Sheets("g_M").Cells(19, 7) = Sheets("d_M").Cells(i_Row, 1)
'Sheets("g_M").Cells(19, 5) = Sheets("d_M").Cells(2, i_Col)



Debug.Print "耗费时间: " & Timer - sngStart

End Sub


