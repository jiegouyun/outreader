Attribute VB_Name = "PKPM_WDISP"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/11
'更新内容:
'1.添加range是否为空判断,解决缺少数据时with报错问题

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'更新内容:
'1.添加数据格式代码,解决with报错问题

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/04/18
'更新内容:
'1.位移比取最值前添加format限制代码

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15
'1.修正高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29
'1.添加高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/29 10:13

'更新内容：
'1.修改最大层间位移角取值范围，只限于地震与风，不含偏心

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/18/ 17:22
'更新内容:
'1.find函数改为精确查找，在末尾添加语句 lookat:=1


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/15/ 17:22
'更新内容:
'1.针对新的general表格更新内容。
'2.最大层间位移角按各工况提取


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/3/ 17:22
'更新内容:
'1.修正最大层间位移角取值范围，包括风工况

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/27/ 21:57
'更新内容:
'1.精确定位层间位移角和层间位移所在列
'2.精确定位最大位移比所在楼层


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/23/ 23:57
'更新内容:
'1.精确定位层间位移角和层间位移所在列
'2.精确定位最大位移比所在楼层


'///////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/20 22:40
'更新内容
'1.解决少列数据情况下的读取

'///////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            PKPM_WDISP.OUT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_PKPM_WDISP(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename  As String, filepath  As String, inputstring  As String


'定义data为读入行的字符串
Dim data As String


'定义循环变量
Dim i, j As Integer

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
Keyword_Story_Drift_EX = "X 方向地震作用下的楼层最大位移"

Keyword_Story_Drift_EXP = "X+ 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_EXN = "X- 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_EY = "Y 方向地震作用下的楼层最大位移"

Keyword_Story_Drift_EYP = "Y+ 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_EYN = "Y- 偶然偏心地震作用下的楼层最大位移"

Keyword_Story_Drift_WX = "X 方向风荷载作用下的楼层最大位移"

Keyword_Story_Drift_WY = "Y 方向风荷载作用下的楼层最大位移"



'位移比
Dim Keyword_Disp_Ratio_FEX As String

Dim Keyword_Disp_Ratio_FEXP As String

Dim Keyword_Disp_Ratio_FEXN As String

Dim Keyword_Disp_Ratio_FEY As String

Dim Keyword_Disp_Ratio_FEYP As String

Dim Keyword_Disp_Ratio_FEYN As String

'赋值
Keyword_Disp_Ratio_FEX = "X 方向地震作用规定水平力下的楼层最大位移"

Keyword_Disp_Ratio_FEXP = "X+偶然偏心地震作用规定水平力下的楼层最大位移"

Keyword_Disp_Ratio_FEXN = "X-偶然偏心地震作用规定水平力下的楼层最大位移"

Keyword_Disp_Ratio_FEY = "Y 方向地震作用规定水平力下的楼层最大位移"

Keyword_Disp_Ratio_FEYP = "Y+偶然偏心地震作用规定水平力下的楼层最大位移"

Keyword_Disp_Ratio_FEYN = "Y-偶然偏心地震作用规定水平力下的楼层最大位移"


'==========================================================================================定义首字符变量


'层间位移角/位移比
Dim Firststring_Disp As String


'=============================================================================================================================生成文件读取路径

'指定文件名为WDISP.out
Filename = "WDISP.OUT"

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件WDISP.out"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    '记录行数
'    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符
   
    '层间位移角/位移比
    Firststring_Disp = Mid(data, 17, 25)

    
    '-------------------------------------------------------------------------------------------读取层间位移角/位移比


    Select Case Firststring_Disp
    
    '----------------------------------------------------------读取X 方向地震作用下的位移、层间位移角
    Case Keyword_Story_Drift_EX
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 18) = StringfromStringforReg(data, "\b\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 26) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
    '----------------------------------------------------------读取X+ 偶然偏心地震作用下的位移、层间位移角
        Case Keyword_Story_Drift_EXP
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 19) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 27) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
    '----------------------------------------------------------读取X- 偶然偏心地震作用下的位移、层间位移角
        Case Keyword_Story_Drift_EXN
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 20) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 28) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
        '----------------------------------------------------------读取Y 方向地震作用下的位移、层间位移角
    Case Keyword_Story_Drift_EY
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 22) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 30) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
    '----------------------------------------------------------读取Y+ 偶然偏心地震作用下的位移、层间位移角
        Case Keyword_Story_Drift_EYP
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 23) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 31) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
    '----------------------------------------------------------读取Y- 偶然偏心地震作用下的层间位移角
        Case Keyword_Story_Drift_EYN
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 24) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 32) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
        '----------------------------------------------------------X 方向风荷载作用下的层间位移角
        Case Keyword_Story_Drift_WX
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 21) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 29) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
        '----------------------------------------------------------Y 方向风荷载作用下的层间位移角
        Case Keyword_Story_Drift_WY
    
        Do While Not EOF(1)
            Line Input #i, data
            
            If CheckRegExpfromString(data, "层间位移角:") Then
                '退出
                Exit Do
            End If
            
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
            
                '结果文件中数据对应层号从大至小，统一为从小到大排列
                
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                
                '读取楼层
                'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                
                '写入平均位移
                Sheets("d_P").Cells(j, 25) = StringfromStringforReg(data, "\d*\.\d*", 2)
                
                '读取下一行中层间位移角数据
                Line Input #i, data
                '写入层间位移角
                Sheets("d_P").Cells(j, 33) = StringfromStringforReg(data, "\b\d+\.\s", 1)
            
            End If
            
        Loop
        
        
        '----------------------------------------------------------读取X 方向地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEX
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值:") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 34) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 40) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
            
        '----------------------------------------------------------读取X+偶然偏心地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEXP
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值:") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 35) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 41) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
        
        '----------------------------------------------------------读取X-偶然偏心地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEXN
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值:") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 36) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 42) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
            
        '----------------------------------------------------------读取Y 方向地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEY
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值:") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 37) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 43) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
            
        '----------------------------------------------------------读取Y+偶然偏心地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEYP
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值:") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 38) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 44) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop
        
        '----------------------------------------------------------读取Y-偶然偏心地震作用规定水平力下的位移比、层间位移比
        Case Keyword_Disp_Ratio_FEYN
        
            Do While Not EOF(1)
                Line Input #i, data
                
                If CheckRegExpfromString(data, "最大位移与层平均位移的比值") Or CheckRegExpfromString(data, "层间位移角:") Then
                    '退出
                    Exit Do
                End If
                
                '如果接连两个数，认为是数据行
                If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") Then
                
                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2
                    
                    '读取楼层
                    'Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\b\d*\b", 1)
                    
                    '写入位移比
                    Sheets("d_P").Cells(j, 39) = StringfromStringforReg(data, "\d*\.\d*", 3)
                    
                    '读取下一行中层间位移比数据
                    Line Input #i, data
                    '写入层间位移比
                    Sheets("d_P").Cells(j, 45) = StringfromStringforReg(data, "\d*\.\d*", 3)
                
                End If
                
            Loop

           
            Debug.Print "读取地震作用"
            Debug.Print "用时: " & Timer - sngStart
            Debug.Print "……"

        End Select
    
        
Loop

'关闭结果文件WDISP.OUT
Close #i

'-------------------------------------------------------------------------------------------高亮最值

'Sheets("d_P").Cells.EntireColumn.AutoFit
'
'Num_All = Sheets("d_P").range("a65536").End(xlUp)
'
'
'Dim ii As Integer
'Dim i_RowID As Integer
'Dim i_Rng As range
'
'
''---------------------------------------------------------位移角
'For ii = 26 To 33
'Dim RT As range
'Set RT = Worksheets("d_P").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(RT, "min", "d_P!R3C" & CStr(ii) & ":R" & CStr(Num_All + 2) & "C" & CStr(ii))
'Next
'
''---------------------------------------------------------位移比
'For ii = 34 To 45
'Set RT = Worksheets("d_P").range(Cells(3, ii), Cells(Num_All + 2, ii))
'Call maxormin(RT, "max", "d_P!R3C" & CStr(ii) & ":R" & CStr(Num_All + 2) & "C" & CStr(ii))
'Next
    
'Sheets("g_P").Cells.EntireColumn.AutoFit
'Sheets("d_P").Cells.EntireColumn.AutoFit
'Sheets("d_P").Cells.NumberFormatLocal = "G/通用格式"


'-------------------------------------------------------------------------------------------读取最大层间位移角、所在楼层及工况
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  风荷载
Sheets("g_P").Cells(10, 5).Formula = "=1&"" / ""&MIN(d_P!AC:AC)"
Sheets("g_P").Cells(10, 7).Formula = "=1&"" / ""&MIN(d_P!AG:AG)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震
Sheets("g_P").Cells(11, 5).Formula = "=1&"" / ""&MIN(d_P!Z:Z)"
Sheets("g_P").Cells(11, 7).Formula = "=1&"" / ""&MIN(d_P!AD:AD)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震+
Sheets("g_P").Cells(12, 5).Formula = "=1&"" / ""&MIN(d_P!AA:AA)"
Sheets("g_P").Cells(12, 7).Formula = "=1&"" / ""&MIN(d_P!AE:AE)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震-
Sheets("g_P").Cells(13, 5).Formula = "=1&"" / ""&MIN(d_P!AB:AB)"
Sheets("g_P").Cells(13, 7).Formula = "=1&"" / ""&MIN(d_P!AF:AF)"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  层间位移角
Sheets("g_P").Cells(14, 4).Formula = "=1&"" / ""&MIN(d_P!Z:Z,d_P!AC:AD,d_P!AG:AG)"


'//////////////////////////////////////////////////////////////////////////////////////////////
'调用EXCEl公式，计算慢，每次需先清除原先的公式才行
'Sheets("g_P").Cells(18, 14).FormulaArray = "=INDEX(d_P!C[-13],SMALL(IF(d_P!C[12]:C[14]=MIN(d_P!C[12]:C[14]),ROW(R[-17]:R[4982]),4^8),ROW(R[-17])))&"""""
'Sheets("g_P").Cells(18, 15).FormulaArray = "=INDEX(d_P!C[-14],SMALL(IF(d_P!C[14]:C[16]=MIN(d_P!C[14]:C[16]),ROW(R[-17]:R[4982]),4^8),ROW(R[-17])))&"""""
'//////////////////////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////////////////////
'使用VBA查询功能，快速


'定义最大层间位移角查询区域
Set iRng_X = Application.Union(range("d_P!Z:Z"), range("d_P!AC:AD"), range("d_P!AG:AG"))

'定义查询变量
Dim i_Min As Double, i_Row As Integer, i_Col As Integer
Dim i_Temp As range

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大层间位移角（实际是查询层间位移角的最小分母）
i_Min = WorksheetFunction.Min(iRng_X)
'i_Min = Format(i_Min, "0000")

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大层间位移角所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大层间位移角所在层，及其工况
    Sheets("g_P").Cells(15, 7) = Sheets("d_P").Cells(i_Row, 1)
    Sheets("g_P").Cells(15, 5) = Sheets("d_P").Cells(2, i_Col)
End If

'返回最大层间位移角所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大层间位移角所在层，及其工况
'Sheets("g_P").Cells(15, 7) = Sheets("d_P").Cells(i_Row, 1)
'Sheets("g_P").Cells(15, 5) = Sheets("d_P").Cells(2, i_Col)

Sheets("d_P").Columns("AH:AS").NumberFormatLocal = "0.00"
'-------------------------------------------------------------------------------------------读取最大位移比
Sheets("g_P").Cells(16, 4).Formula = "=MAX(d_P!AH:AM)"

'------------------------------------------------------------------------------------------- 查询最大位移比所在楼层

'定义最大位移比查询区域
Set iRng_X = range("d_P!AH:AM")

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大位移比
i_Min = Application.WorksheetFunction.Max(range("d_P!AH:AM"))
i_Min = Format(i_Min, "0.00")

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大位移比所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大位移比所在层，及其工况
    Sheets("g_P").Cells(17, 7) = Sheets("d_P").Cells(i_Row, 1)
    Sheets("g_P").Cells(17, 5) = Sheets("d_P").Cells(2, i_Col)
End If

'返回最大位移比所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大位移比所在层，及其工况
'Sheets("g_P").Cells(17, 7) = Sheets("d_P").Cells(i_Row, 1)
'Sheets("g_P").Cells(17, 5) = Sheets("d_P").Cells(2, i_Col)


'-------------------------------------------------------------------------------------------读取最大层间位移比
Sheets("g_P").Cells(18, 4).Formula = "=MAX(d_P!AN:AS)"

'------------------------------------------------------------------------------------------- 查询最大层间位移比所在楼层

'定义最大位移比查询区域
Set iRng_X = range("d_P!AN:AS")

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
'查询区域内最大位移比
i_Min = WorksheetFunction.Max(range("d_P!AN:AS"))
i_Min = Format(i_Min, "0.00")

Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

If Not i_Temp Is Nothing Then
    '返回最大位移比所在行号、列号
    i_Row = i_Temp.Row
    i_Col = i_Temp.column
    '返回最大位移比所在层，及其工况
    Sheets("g_P").Cells(19, 7) = Sheets("d_P").Cells(i_Row, 1)
    Sheets("g_P").Cells(19, 5) = Sheets("d_P").Cells(2, i_Col)
End If

'返回最大位移比所在行号、列号
'i_Row = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).Row
'i_Col = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=1).column
'返回最大位移比所在层，及其工况
'Sheets("g_P").Cells(19, 7) = Sheets("d_P").Cells(i_Row, 1)
'Sheets("g_P").Cells(19, 5) = Sheets("d_P").Cells(2, i_Col)

Debug.Print "耗费时间: " & Timer - sngStart

End Sub



