Attribute VB_Name = "PKPM_WPJ_UC"
Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/18 22:05

'更新内容：
'1.find添加精确查找参数，lookat:=1

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/3 20:05

'更新内容：
'1.Colum与Wall表格改名为CR_P与WR_P；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/7

'更新内容：
'1.增加range为空时的判断，解决某层没有墙柱引起的函数错误；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/4 想起去年去香港

'更新内容：
'1.数据写入行列互换，相应调整代码；
'2.对最大轴压比数据添加底色，方便查看。
'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/3

'更新内容：
'1.添加说明；
'2.寻找Range最大值的代码换为函数表达。

'////////////////////////////////////////////////////////////////////////////////////////////

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            PKPM_WPJ_UC.OUT部分代码                   ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

Sub OUTReader_PKPM_WPJ_UC(Path As String, num As Integer)

'==========================================================================================写入层号
Sheets("CR_P").Cells(num + 1, 1) = CStr(num) & "F"
Sheets("WR_P").Cells(num + 1, 1) = CStr(num) & "F"



'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring  As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer

'定义最值列数索引变量
Dim C_C, C_W As Integer

'==========================================================================================定义关键词变量

'柱、墙
Dim Keyword_Column, Keyword_Wall As String
'赋值
Keyword_Column = "柱配筋和验算输出"
Keyword_Wall = "墙配筋和验算输出"

'柱、墙轴压比
Dim Keyword_Column_UC, Keyword_Wall_UC As String
'赋值
Keyword_Column_UC = "Uc="
Keyword_Wall_UC = "Uc="


'==========================================================================================定义首字符变量

'柱、墙
Dim FirstString_Column, FirstString_Wall As String
'柱、墙轴压比
Dim FirstString_Column_UC, FirstString_Wall_UC As String


'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "WPJ" & CStr(num) & ".OUT"

'生成完整文件路径
filepath = Path & "\" & Filename
'Debug.Print path
'Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #1



'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wpj" & CStr(num); ".out; "
Debug.Print "读取相关指标"
Debug.Print "……"





Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
    
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Column = Mid(data, 33, 8)
    FirstString_Wall = Mid(data, 33, 8)

    '--------------------------------------------------------------------------读取柱的轴压比
    i = 0
    If FirstString_Column = Keyword_Column Then
        Debug.Print "读取" & CStr(num) & "层柱的轴压比……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Column_UC = Mid(data, 20, 3)
            If FirstString_Column_UC = Keyword_Column_UC Then
                Sheets("CR_P").Cells(num + 1, 2 + i) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                Sheets("CR_P").Cells(1, 2 + i) = i + 1
                i = i + 1
            End If
            If CheckRegExpfromString(data, "===") = True Then
                Exit Do
            End If
        Loop
    End If
    '--------------------------------------------------------------------------读取墙的轴压比
    i = 0
    If FirstString_Wall = Keyword_Wall Then
    Debug.Print "读取" & CStr(num) & "层墙的轴压比……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Wall_UC = Mid(data, 20, 3)
            If FirstString_Wall_UC = Keyword_Wall_UC Then
                Sheets("WR_P").Cells(num + 1, 2 + i) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                Sheets("WR_P").Cells(1, 2 + i) = i + 1
                i = i + 1
            End If
            If CheckRegExpfromString(data, "===") = True Then
                Exit Do
            End If
        Loop
    End If
Loop

Close #1





'===========================================================================================读取Num层最大轴压比及其构件编号并写入dis
'--------------------------------------------------------------------------柱
If Sheets("CR_P").Cells(num + 1, 2) = "" Then
    Sheets("d_P").Cells(num + 2, 56) = 0
    Sheets("d_P").Cells(num + 2, 57) = 0
Else
'最大轴压比所在列数
C_C = IndexMaxofRange(Sheets("CR_P").range(Sheets("CR_P").Cells(num + 1, 2), Sheets("CR_P").Cells(num + 1, 3000)))(2)
'将最大轴压比及构件编号写入dis
Sheets("d_P").Cells(num + 2, 56) = Worksheets("CR_P").Cells(num + 1, C_C)
Sheets("d_P").Cells(num + 2, 57) = C_C - 1
Worksheets("CR_P").Cells(num + 1, C_C).Interior.ColorIndex = 4
End If
'--------------------------------------------------------------------------墙
If Sheets("WR_P").Cells(num + 1, 2) = "" Then
    Sheets("d_P").Cells(num + 2, 58) = 0
    Sheets("d_P").Cells(num + 2, 59) = 0
Else
'最大轴压比所在列数
C_W = IndexMaxofRange(Sheets("WR_P").range(Sheets("WR_P").Cells(num + 1, 2), Sheets("WR_P").Cells(num + 1, 3000)))(2)
'将最大轴压比及构件编号写入dis
Sheets("d_P").Cells(num + 2, 58) = Worksheets("WR_P").Cells(num + 1, C_W)
Sheets("d_P").Cells(num + 2, 59) = C_W - 1
Worksheets("WR_P").Cells(num + 1, C_W).Interior.ColorIndex = 4

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

