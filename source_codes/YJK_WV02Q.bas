Attribute VB_Name = "YJK_WV02Q"
'该模块读取指定文件路径下的WV02Q文件，在Distribution工作表里输出剪力调整情况

'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/3/19
'1.修正读取剪力比的错误


'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/3/16
'1.数据读取有漏，修改代码；


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/18
'1.添加读取倾覆力矩百分比代码


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/18
'1.更正柱剪力百分比的读取位置

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/4

'更新内容：
'1.改正剪力调整写入bug;

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/3

'更新内容：
'1.更改代码，除去对全局变量的依赖；

'////////////////////////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            YJK_WV02Q.OUT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub OUTReader_YJK_WV02Q(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring  As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim c_i, c_j As Integer
Dim i As Integer


'==========================================================================================定义关键词变量
'柱剪力
Dim Keyword_VC As String
Keyword_VC = "框架柱地震剪力百分比"

Dim Keyword_VCX As String
Keyword_VCX = "X"

Dim Keyword_VCY As String
Keyword_VCY = "Y"

'剪力调整系数
Dim Keyword_VT As String
Keyword_VT = "0.2V0"

Dim Keyword_V As String
Keyword_V = "Coef_x"

'==========================================================================================定义首字符变量
'柱剪力
Dim FirstString_VC As String
Dim FirstString_VCX As String
Dim FirstString_VCY As String
'剪力调整系数
Dim FirstString_V As String
Dim FirstString_VT As String


'==========================================================================================生成文件读取路径

'指定文件名
Filename = "WV02Q.OUT"

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #1

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件WV02Q.OUT："
Debug.Print "读取相关数据"
Debug.Print "……"

'初始化剪力调整系数循环变量
i = 0

Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    FirstString_VC = Mid(data, 23, 10)
    FirstString_VT = Mid(data, 28, 5)
    FirstString_MRK = Mid(data, 11, 22)
    'FirstString_MRZ = Mid(data, 16, 27)
   
   
    '--------------------------------------------------------------------------读取柱剪力及其所占总剪力的百分比
    'c_i = 0
    'c_j = 0
    '寻找楼层数
    Dim NA As Integer
    If FirstString_VC = Keyword_VC Then
        Line Input #1, data
        Debug.Print "读取柱剪力及其所占总剪力的百分比……"
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_VCX = Mid(data, 16, 1)
            FirstString_VCY = Mid(data, 16, 1)
            If FirstString_VCX = Keyword_VCX Then
                 'Debug.Print "柱剪力test"
                 NA = extractNumberFromString(data, 1)
                 Sheets("d_Y").Cells(NA + 2, 48) = StringfromStringforReg(data, "\S+", 4)
                 Sheets("d_Y").Cells(NA + 2, 49) = StringfromStringforReg(data, "\S+", 8)
                 'c_i = c_i + 1
            End If
            If FirstString_VCY = Keyword_VCY Then
                NA = extractNumberFromString(data, 1)
                 Sheets("d_Y").Cells(NA + 2, 51) = StringfromStringforReg(data, "\S+", 4)
                 Sheets("d_Y").Cells(NA + 2, 52) = StringfromStringforReg(data, "\S+", 8)
                 'c_j = c_j + 1
            End If
            If CheckRegExpfromString(data, "\*") Then
                'Debug.Print "柱剪力终止test"
                Exit Do
            End If
        Loop
    End If

    '--------------------------------------------------------------------------读取X/Y向剪力调整系数
    '寻找层数
    Dim NA2 As Integer
    NA2 = 1
    If FirstString_VT = Keyword_VT Then
        'NA2 = extractNumberFromString(data, 1)
        Debug.Print "读取柱剪力调整系数……"
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_V = Mid(data, 10, 6)
            If FirstString_V = Keyword_V Then
                 'Debug.Print "剪力调整test"
                 Line Input #1, data
                 'Line Input #1, data
                 Sheets("d_Y").Cells(NA2 + 2, 50) = StringfromStringforReg(data, "\S+", 1)
                 Sheets("d_Y").Cells(NA2 + 2, 53) = StringfromStringforReg(data, "\S+", 2)
                 NA2 = NA2 + 1
            End If
            'If CheckRegExpfromString(data, "==") Then
                'Debug.Print "剪力调整终止test"
                'Exit Do
            'End If
        Loop
    End If
    
    '--------------------------------------------------------------------------读取X/Y向倾覆力矩百分比(抗规)
    If FirstString_MRK = "规定水平力下框架柱、短肢墙地震倾覆弯矩百分比" Then
        Debug.Print "读取百分比（K）……"
        Line Input #1, data
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_VCX = Mid(data, 16, 1)
            FirstString_VCY = Mid(data, 16, 1)
            If FirstString_VCX = "X" Then
                 'Debug.Print "柱剪力test"
                 NA = extractNumberFromString(data, 1)
                 'Debug.Print NA
                 Sheets("d_Y").Cells(NA + 2, 70) = StringfromStringforReg(data, "\S+", 4)
                 Sheets("d_Y").Cells(NA + 2, 72) = StringfromStringforReg(data, "\S+", 5)
                 'c_i = c_i + 1
            End If
            If FirstString_VCY = "Y" Then
                NA = extractNumberFromString(data, 1)
                 Sheets("d_Y").Cells(NA + 2, 71) = StringfromStringforReg(data, "\S+", 4)
                 Sheets("d_Y").Cells(NA + 2, 73) = StringfromStringforReg(data, "\S+", 5)
                 'c_j = c_j + 1
            End If
            If CheckRegExpfromString(data, "\*") Then
                'Debug.Print "柱剪力终止test"
                Exit Do
            End If
        Loop
    End If
Loop

Close #1
Sheets("g_Y").Cells(53, 5).Formula = "=d_Y!" & "BR" & Num_Base + 3
Sheets("g_Y").Cells(53, 7).Formula = "无"
Sheets("g_Y").Cells(54, 5).Formula = "=d_Y!" & "BS" & Num_Base + 3
Sheets("g_Y").Cells(54, 7).Formula = "无"

Debug.Print "读取柱剪力信息耗费时间: " & Timer - sngStart


End Sub
