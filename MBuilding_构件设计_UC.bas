Attribute VB_Name = "MBuilding_构件设计_UC"
Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。


'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间：2013/07/28 11：30
'更形内容：
'1.添加柱轴压比提取

'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间：2013/07/20 12：10
'更形内容：
'1.移植PKPM_WPJ_UC，提供的轴压比数据文件中只有墙柱轴压比信息，没有柱轴压比信息，未能验证提取柱轴压比信息是否正确



'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****             MBuilding_构件设计及验算结果.TXT部分代码                 ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

Sub OUTReader_MBuilding_构件设计_UC(Path As String, num)

'==========================================================================================写入层号
Sheets("CR_M").Cells(num + 1, 1) = CStr(num) & "F"
Sheets("WR_M").Cells(num + 1, 1) = CStr(num) & "F"



'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径
Dim Filename, filepath As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer

'定义最值列数索引变量
Dim C_C, C_W As Integer

'==========================================================================================定义关键词变量

'柱、墙
Dim Keyword_Column, Keyword_WC As String
'赋值
Keyword_Column = "钢筋混凝土柱配筋和设计结果"
Keyword_WC = "钢筋混凝土墙柱配筋和设计结果"

'柱、墙轴压比
Dim Keyword_Column_UC, Keyword_WC_UC As String
'赋值
Keyword_Column_UC = "NAF ="
Keyword_WC_UC = "NAF ="


'==========================================================================================定义首字符变量

'柱、墙
Dim FirstString_Column, FirstString_WC As String
'柱、墙轴压比
Dim FirstString_Column_UC, FirstString_WC_UC As String


'==========================================================================================生成文件读取路径
'指定文件名为 各构件设计及验算结果_Num.txt
Filename = Dir(Path & "\*_各构件设计及验算结果_" & CStr(num) & "F.txt")

'生成完整文件路径
filepath = Path & "\" & Filename
'Debug.Print path
'Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #1



'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件：构件设计及验算结果_" & CStr(num); "F.txt; "
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, data '读文本文件一行
    
    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Column = Mid(data, 8 + Len(num), 13)
    FirstString_WC = Mid(data, 8 + Len(num), 14)
'   Debug.Print num
    '--------------------------------------------------------------------------读取柱的轴压比
    i = 0
    If FirstString_Column = Keyword_Column Then
        Debug.Print "读取" & CStr(num) & "层柱的轴压比……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Column_UC = Mid(data, 29, 5)
            Debug.Print FirstString_Column_UC
            If FirstString_Column_UC = Keyword_Column_UC Then
            Debug.Print "XX1"
                Sheets("CR_M").Cells(num + 1, 2 + i) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                Sheets("CR_M").Cells(1, 2 + i) = i + 1
                i = i + 1
            End If
            If CheckRegExpfromString(data, "===") Then
                Exit Do
            End If
        Loop
    End If
    '--------------------------------------------------------------------------读取墙柱的轴压比
    i = 0
    If FirstString_WC = Keyword_WC Then
    Debug.Print "读取" & CStr(num) & "层墙柱的轴压比……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_WC_UC = Mid(data, 19, 5)
            Debug.Print FirstString_WC_UC
            If FirstString_WC_UC = Keyword_WC_UC Then
            Debug.Print "XX2"
                Sheets("WR_M").Cells(num + 1, 2 + i) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                Sheets("WR_M").Cells(1, 2 + i) = i + 1
                i = i + 1
            End If
            If CheckRegExpfromString(data, "===") Then
                Exit Do
            End If
        Loop
    End If
Loop

Close #1



'===========================================================================================读取Num层最大轴压比及其构件编号并写入dis
'--------------------------------------------------------------------------柱
If Sheets("CR_M").Cells(num + 1, 2) = "" Then
    Sheets("d_M").Cells(num + 2, 56) = 0
    Sheets("d_M").Cells(num + 2, 57) = 0
Else
'最大轴压比所在列数
C_C = IndexMaxofRange(Sheets("CR_M").range(Sheets("CR_M").Cells(num + 1, 2), Sheets("CR_M").Cells(num + 1, 3000)))(2)
'将最大轴压比及构件编号写入dis
Sheets("d_M").Cells(num + 2, 56) = Worksheets("CR_M").Cells(num + 1, C_C)
Sheets("d_M").Cells(num + 2, 57) = C_C - 1
Worksheets("CR_M").Cells(num + 1, C_C).Interior.ColorIndex = 4
End If
'--------------------------------------------------------------------------墙柱
If Sheets("WR_M").Cells(num + 1, 2) = "" Then
    Sheets("d_M").Cells(num + 2, 58) = 0
    Sheets("d_M").Cells(num + 2, 59) = 0
Else
'最大轴压比所在列数
C_W = IndexMaxofRange(Sheets("WR_M").range(Sheets("WR_M").Cells(num + 1, 2), Sheets("WR_M").Cells(num + 1, 3000)))(2)
'将最大轴压比及构件编号写入dis
Sheets("d_M").Cells(num + 2, 58) = Worksheets("WR_M").Cells(num + 1, C_W)
Sheets("d_M").Cells(num + 2, 59) = C_W - 1
Worksheets("WR_M").Cells(num + 1, C_W).Interior.ColorIndex = 4

End If


Debug.Print "读取" & CStr(num) & "层墙柱轴压比耗费时间: " & Timer - sngStart


End Sub

Function IndexMaxofRange(index_Range As range)
Dim Max, R, C As Integer
Max = WorksheetFunction.Max(index_Range)
R = index_Range.Find(Max, After:=index_Range.Cells(index_Range.Rows.Count, index_Range.Columns.Count), LookIn:=xlValues, lookat:=1).Row
C = index_Range.Find(Max, After:=index_Range.Cells(index_Range.Rows.Count, index_Range.Columns.Count), LookIn:=xlValues, lookat:=1).column
IndexMaxofRange = Array(Max, R, C)
End Function

