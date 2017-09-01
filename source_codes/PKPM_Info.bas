Attribute VB_Name = "PKPM_Info"
Option Explicit
Public i_c1, i_w1, i_b1, i_wb1, i_c2, i_w2, i_b2, i_wb2 As Integer

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            提取构件配筋信息代码                      ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.图表高宽改为外部参数输入，方便修改

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/28/ 14:30
'更新内容:
'1.计算比值时先判断分母是否为零，解决由于构件信息不完整造成的计算溢出问题
'2.将所有比值大于一的单元格高亮

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/16/ 9:06
'更新内容:
'1.将表格生成和图表生成分开
'2.表名简化,如ColumnInfo改为CI

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12/ 21:26
'更新内容:
'1.添加墙梁配筋读取
'2.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////


'更新时间:2013/7/29/ 21:26
'更新内容:
'1.添加墙暗柱配筋率的自动计算
'2.添加梁最大配筋率的自动取大值





'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub OUTReader_PKPM_CompareS_table(path1 As String, path2 As String, startf As Integer, endf As Integer)

'======================================================================================================生成表格
Call AddInfoSheet

i_c1 = 0
i_w1 = 0
i_b1 = 0
i_wb1 = 0
i_c2 = 0
i_w2 = 0
i_b2 = 0
i_wb2 = 0


Dim i As Integer

For i = startf To endf

Call OUTReader_PKPM_Info(path1, path2, i)

Next

'柱配筋比值计算
For i = 1 To i_c1
With Sheets("CI")
    '如果分母不为零，则计算比值
    If Not .Cells(3 + i, 8) = 0 Then .Cells(3 + i, 16) = .Cells(3 + i, 9) / .Cells(3 + i, 8)
    '如果比值大于零，则高亮显示
    If .Cells(3 + i, 16) > 1 Then .Cells(3 + i, 16).Interior.ColorIndex = 4
    If Not .Cells(3 + i, 10) = 0 Then .Cells(3 + i, 17) = .Cells(3 + i, 11) / .Cells(3 + i, 10)
    If .Cells(3 + i, 17) > 1 Then .Cells(3 + i, 17).Interior.ColorIndex = 4
End With
Next
'墙配筋比值计算
For i = 1 To i_w1
With Sheets("WCI")
    If Not .Cells(3 + i, 9) = 0 Then .Cells(3 + i, 19) = .Cells(3 + i, 10) / .Cells(3 + i, 9)
    If .Cells(3 + i, 19) > 1 Then .Cells(3 + i, 19).Interior.ColorIndex = 4
    If Not .Cells(3 + i, 13) = 0 Then .Cells(3 + i, 20) = .Cells(3 + i, 14) / .Cells(3 + i, 13)
    If .Cells(3 + i, 20) > 1 Then .Cells(3 + i, 20).Interior.ColorIndex = 4
End With
Next

'Dim as_b1, as_b2, as_b3, as_b4

'梁配筋比值计算
For i = 1 To i_b1

With Sheets("BI")
    '若梁顶筋为0，则取底筋的1/4
    If .Cells(3 + i, 6) = 0 Then
        .Cells(3 + i, 6) = 0.25 * .Cells(3 + i, 7)
    End If
    If .Cells(3 + i, 8) = 0 Then
        .Cells(3 + i, 8) = 0.25 * .Cells(3 + i, 7)
    End If
    If .Cells(3 + i, 9) = 0 Then
        .Cells(3 + i, 9) = 0.25 * .Cells(3 + i, 10)
    End If
    If .Cells(3 + i, 11) = 0 Then
        .Cells(3 + i, 11) = 0.25 * .Cells(3 + i, 10)
    End If
    If Not .Cells(3 + i, 6) = 0 Then .Cells(3 + i, 14) = .Cells(3 + i, 9) / .Cells(3 + i, 6)
    If .Cells(3 + i, 14) > 1 Then .Cells(3 + i, 14).Interior.ColorIndex = 4
    If Not .Cells(3 + i, 7) = 0 Then .Cells(3 + i, 15) = .Cells(3 + i, 10) / .Cells(3 + i, 7)
    If .Cells(3 + i, 15) > 1 Then .Cells(3 + i, 15).Interior.ColorIndex = 4
    If Not .Cells(3 + i, 8) = 0 Then .Cells(3 + i, 16) = .Cells(3 + i, 11) / .Cells(3 + i, 8)
    If .Cells(3 + i, 16) > 1 Then .Cells(3 + i, 16).Interior.ColorIndex = 4
    If Not .Cells(3 + i, 12) = 0 Then .Cells(3 + i, 17) = .Cells(3 + i, 13) / .Cells(3 + i, 12)
    If .Cells(3 + i, 17) > 1 Then .Cells(3 + i, 17).Interior.ColorIndex = 4
End With
Next

'墙梁配筋比值计算
For i = 1 To i_wb1

With Sheets("WBI")
    '若梁顶筋为0，则取底筋的1/4
    If .Cells(3 + i, 6) = 0 Then
        .Cells(3 + i, 6) = 0.25 * .Cells(3 + i, 7)
    End If
    If .Cells(3 + i, 8) = 0 Then
        .Cells(3 + i, 8) = 0.25 * .Cells(3 + i, 7)
    End If
    If .Cells(3 + i, 9) = 0 Then
        .Cells(3 + i, 9) = 0.25 * .Cells(3 + i, 10)
    End If
    If .Cells(3 + i, 11) = 0 Then
        .Cells(3 + i, 11) = 0.25 * .Cells(3 + i, 10)
    End If
    .Cells(3 + i, 14) = .Cells(3 + i, 9) / .Cells(3 + i, 6)
    .Cells(3 + i, 15) = .Cells(3 + i, 10) / .Cells(3 + i, 7)
    .Cells(3 + i, 16) = .Cells(3 + i, 11) / .Cells(3 + i, 8)
    .Cells(3 + i, 17) = .Cells(3 + i, 13) / .Cells(3 + i, 12)
End With
Next


End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub OUTReader_PKPM_CompareS_figure()

'======================================================================================================生成图表

Dim sh As Worksheet
'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = "figure_Info" Then
        sh.Delete
    End If
Next

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284

Call Addsh("figure_Info")

Call add_chart_2("CI", "P4:P" & i_c1 + 3, "Q4:Q" & i_c2 + 3, "A4:A" & i_c1 + 3, "主筋配筋率", "配箍率", "中震/小震", "柱编号", 0 * Width, 0 * Hight, Width, Hight)

Call add_chart_2("WCI", "S4:S" & i_w1 + 3, "T4:T" & i_w2 + 3, "A4:A" & i_w1 + 3, "暗柱配筋", "分布钢筋", "中震/小震", "墙编号", 1 * Width, 0 * Hight, Width, Hight)

Call add_chart_4("BI", "N4:N" & i_b1 + 3, "O4:O" & i_b2 + 3, "P4:P" & i_b1 + 3, "Q4:Q" & i_b2 + 3, "A4:A" & i_b1 + 3, "I点顶筋", "中点底筋", "J点顶筋", "箍筋", "中震/小震", "梁编号", 2 * Width, 0 * Hight, Width, Hight)

Call add_chart_4("WBI", "N4:N" & i_wb1 + 3, "O4:O" & i_wb2 + 3, "P4:P" & i_wb1 + 3, "Q4:Q" & i_wb2 + 3, "A4:A" & i_wb1 + 3, "I点顶筋", "中点底筋", "J点顶筋", "箍筋", "中震/小震", "墙梁编号", 3 * Width, 0 * Hight, Width, Hight)

Sheets("figure_Info").Select
'Sheets("BI").range("L4").Select
'Selection.AutoFill Destination:=range("L4:L" & i_b1), Type:=xlFillDefault
'Sheets("BI").range("M4").Select
'Selection.AutoFill Destination:=range("M4:M" & i_b2), Type:=xlFillDefault


End Sub


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub OUTReader_PKPM_Info(path1 As String, path2 As String, num As Integer)


'==========================================================================================写入层号

'Sheets("WCI").Cells(Num + 1, 2) = CStr(Num) & "F"
'Sheets("CI").Cells(Num + 1, 2) = CStr(Num) & "F"
'Sheets("BI").Cells(Num + 1, 2) = CStr(Num) & "F"
Dim n As Integer
n = num


'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, filepath2, inputstring  As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义最值列数索引变量
Dim C_C, C_W As Integer

'==========================================================================================定义关键词变量

'柱、墙、梁编号行关键词
Dim Keyword_Column, Keyword_Wall, Keyword_Beam, Keyword_WBeam As String
'赋值
Keyword_Column = "N-C="
Keyword_Wall = "N-WC="
Keyword_Beam = "N-B="
Keyword_WBeam = "N-WB="

'柱、墙轴压比行关键词
Dim Keyword_Column_UC, Keyword_Wall_UC As String
'赋值
Keyword_Column_UC = "Uc="
Keyword_Wall_UC = "Uc="

'梁顶筋、底筋，箍筋配筋率关键词
Dim Keyword_Beam_Top, Keyword_Beam_Btm, Keyword_Beam_Rsv As String
Dim Keyword_WBeam_Top, Keyword_WBeam_Btm, Keyword_WBeam_Rsv As String
'赋值
Keyword_Beam_Top = "Top Ast"
Keyword_Beam_Btm = "Btm Ast"
Keyword_Beam_Rsv = "Rsv"
Keyword_WBeam_Top = "Top Ast"
Keyword_WBeam_Btm = "Btm Ast"
Keyword_WBeam_Rsv = "Rsv"


'柱、墙抗剪承载力行关键词
Dim Keyword_Column_V, Keyword_Wall_V As String
'赋值
Keyword_Column_V = "抗剪承载力"
Keyword_Wall_V = "抗剪承载力"

'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_Column, FirstString_Wall, FirstString_Beam, FirstString_WBeam As String
'柱、墙轴压比，梁配筋率
Dim FirstString_Column_UC, FirstString_Wall_UC, FirstString_Beam_S, FirstString_WBeam_S As String


'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "WPJ" & CStr(num) & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename
filepath2 = path2 & "\" & Filename

'Debug.Print path1
'Debug.Print filepath

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i


'===========================================================================================逐行读取文本

Debug.Print "开始遍历小震结果文件wpj" & CStr(num); ".out; "
Debug.Print "读取相关指标"
Debug.Print "……"





Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
    
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Column = Mid(data, 2, 4)
    FirstString_Wall = Mid(data, 2, 5)
    FirstString_Beam = Mid(data, 2, 4)
    FirstString_WBeam = Mid(data, 2, 5)

    '--------------------------------------------------------------------------读取柱的信息
    If FirstString_Column = Keyword_Column Then
        Debug.Print "读取" & CStr(num) & "层柱信息……"
        '写入序号
        Sheets("CI").Cells(i_c1 + 4, 1) = i_c1 + 1
        '写入楼层号
        Sheets("CI").Cells(i_c1 + 4, 2) = n
        '读取柱编号
        Sheets("CI").Cells(i_c1 + 4, 3) = "NC-" & extractNumberFromString(data, 1)
        '读取柱截面
        Sheets("CI").Cells(i_c1 + 4, 4) = extractNumberFromString(data, 3)
        Sheets("CI").Cells(i_c1 + 4, 5) = extractNumberFromString(data, 4)
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Column_UC = Mid(data, 20, 3)
            If FirstString_Column_UC = Keyword_Column_UC Then
                '读取柱轴压比
                Sheets("CI").Cells(i_c1 + 4, 6) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                '读取柱主筋配筋率
                Sheets("CI").Cells(i_c1 + 4, 8) = Mid(data, 34, 6)
                '读取柱箍筋配筋率
                Sheets("CI").Cells(i_c1 + 4, 10) = Mid(data, 48, 6)
            End If
            
            If Mid(data, 2, 5) = Keyword_Column_V Then
                '读取抗剪承载力
                Sheets("CI").Cells(i_c1 + 4, 12) = extractNumberFromString(data, 1)
                Sheets("CI").Cells(i_c1 + 4, 14) = extractNumberFromString(data, 2)
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_c1 = i_c1 + 1
                Exit Do
            End If
            
        Loop
    End If
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        Debug.Print "读取" & CStr(num) & "层墙信息……"
        '写入序号
        Sheets("WCI").Cells(i_w1 + 4, 1) = i_w1 + 1
        '写入楼层号
        Sheets("WCI").Cells(i_w1 + 4, 2) = n
        '读取墙编号
        Sheets("WCI").Cells(i_w1 + 4, 3) = "NWC-" & extractNumberFromString(data, 1)
        '读取墙截面
        Dim B_w As Integer, H_w As Integer
        Sheets("WCI").Cells(i_w1 + 4, 4) = extractNumberFromString(data, 4) * 1000
        B_w = Sheets("WCI").Cells(i_w1 + 4, 4)
        Sheets("WCI").Cells(i_w1 + 4, 5) = extractNumberFromString(data, 5) * 1000
        H_w = Sheets("WCI").Cells(i_w1 + 4, 5)
        Sheets("WCI").Cells(i_w1 + 4, 6) = extractNumberFromString(data, 6) * 1000
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Wall_UC = Mid(data, 20, 3)
            If FirstString_Wall_UC = Keyword_Wall_UC Then
                '读取墙轴压比
                Debug.Print "读取" & CStr(num) & "层墙轴压比……"
                Sheets("WCI").Cells(i_w1 + 4, 7) = StringfromStringforReg(data, "\s+0\.\d*", 1)
            End If
            
            If Mid(data, 7, 2) = "M=" And Mid(data, 19, 2) = "N=" Then
                '读取暗柱配筋面积
                Dim As_az As Single
                Debug.Print "读取" & CStr(num) & "层墙暗柱配筋……"
                Sheets("WCI").Cells(i_w1 + 4, 9) = extractNumberFromString(data, 4)
                As_az = Sheets("WCI").Cells(i_w1 + 4, 9)
                '暗柱配筋若为0，即构造配筋，改为1，方便小中震比值的计算
                If As_az = 0 Then
                    Sheets("WCI").Cells(i_w1 + 4, 9) = 1
                End If
            End If
            
            If Mid(data, 7, 2) = "V=" And Mid(data, 19, 2) = "N=" Then
                '读取水平分布筋配筋面积
                 Debug.Print "读取" & CStr(num) & "层墙水平配筋……"
                Sheets("WCI").Cells(i_w1 + 4, 11) = extractNumberFromString(data, 4)
                Sheets("WCI").Cells(i_w1 + 4, 13) = extractNumberFromString(data, 5)
            End If
            
            If Mid(data, 2, 5) = Keyword_Wall_V Then
                '读取抗剪承载力
                 Debug.Print "读取" & CStr(num) & "层墙抗剪承载力……"
                Sheets("WCI").Cells(i_w1 + 4, 15) = extractNumberFromString(data, 1)
                Sheets("WCI").Cells(i_w1 + 4, 17) = extractNumberFromString(data, 2)
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_w1 = i_w1 + 1
                Exit Do
            End If
        Loop
    End If
    
    '--------------------------------------------------------------------------读取梁信息
    If FirstString_Beam = Keyword_Beam Then
        Debug.Print "读取" & CStr(num) & "层梁信息……"
        '写入序号
        Sheets("BI").Cells(i_b1 + 4, 1) = i_b1 + 1
        '写入楼层号
        Sheets("BI").Cells(i_b1 + 4, 2) = n
        '读取梁编号
        Sheets("BI").Cells(i_b1 + 4, 3) = "NB-" & extractNumberFromString(data, 1)
        '读取梁截面
        Sheets("BI").Cells(i_b1 + 4, 4) = extractNumberFromString(data, 5)
        Sheets("BI").Cells(i_b1 + 4, 5) = extractNumberFromString(data, 6)
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Beam_S = Mid(data, 1, 7)
            If FirstString_Beam_S = Keyword_Beam_Top Then
                '读取梁顶筋
                Do While Not EOF(1)
                    Line Input #1, data
                    FirstString_Beam_S = Mid(data, 1, 7)
                    If FirstString_Beam_S = "% Steel" Then
                        Debug.Print "读取" & CStr(num) & "层梁顶筋配筋率……"
                        Sheets("BI").Cells(i_b1 + 4, 6) = extractNumberFromString(data, 1)
                        Sheets("BI").Cells(i_b1 + 4, 8) = extractNumberFromString(data, 9)
'                        Sheets("BI").Cells(i_b1 + 4, 8) = "=MAX(F4:H4)"
                    End If
                
                    If FirstString_Beam_S = Keyword_Beam_Btm Then
                        Exit Do
                    End If
                Loop
                
            End If
            
            '读取梁底筋
            FirstString_Beam_S = Mid(data, 1, 7)
            If FirstString_Beam_S = "% Steel" Then
                Debug.Print "读取" & CStr(num) & "层梁底筋配筋率……"
                Sheets("BI").Cells(i_b1 + 4, 7) = extractNumberFromString(data, 5)
            End If
            
            
            '读取梁配箍率
            If Mid(data, 2, 3) = Keyword_Beam_Rsv Then
                Debug.Print "读取" & CStr(num) & "层梁配箍率……"
                Sheets("BI").Cells(i_b1 + 4, 12) = extractNumberFromString(data, 1)
                If Sheets("BI").Cells(i_b1 + 4, 12) < extractNumberFromString(data, 5) Then
                    Sheets("BI").Cells(i_b1 + 4, 12) = extractNumberFromString(data, 5)
                End If
                If Sheets("BI").Cells(i_b1 + 4, 12) < extractNumberFromString(data, 9) Then
                    Sheets("BI").Cells(i_b1 + 4, 12) = extractNumberFromString(data, 9)
                End If
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_b1 = i_b1 + 1
                Exit Do
            End If

        Loop
    End If
    
    
    '--------------------------------------------------------------------------读取墙梁信息
    If FirstString_WBeam = Keyword_WBeam Then
        Debug.Print "读取" & CStr(num) & "层墙梁信息……"
        '写入序号
        Sheets("WBI").Cells(i_wb1 + 4, 1) = i_wb1 + 1
        '写入楼层号
        Sheets("WBI").Cells(i_wb1 + 4, 2) = n
        '读取梁编号
        Sheets("WBI").Cells(i_wb1 + 4, 3) = "NWB-" & extractNumberFromString(data, 1)
        '读取梁截面
        Sheets("WBI").Cells(i_wb1 + 4, 4) = extractNumberFromString(data, 5)
        Sheets("WBI").Cells(i_wb1 + 4, 5) = extractNumberFromString(data, 6)
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_WBeam_S = Mid(data, 1, 7)
            If FirstString_WBeam_S = Keyword_WBeam_Top Then
                '读取梁顶筋
                Do While Not EOF(1)
                    Line Input #1, data
                    FirstString_WBeam_S = Mid(data, 1, 7)
                    If FirstString_WBeam_S = "% Steel" Then
                        Debug.Print "读取" & CStr(num) & "层墙梁顶筋配筋率……"
                        Sheets("WBI").Cells(i_wb1 + 4, 6) = extractNumberFromString(data, 1)
                        Sheets("WBI").Cells(i_wb1 + 4, 8) = extractNumberFromString(data, 9)
'                        Sheets("WBI").Cells(i_wb1 + 4, 8) = "=MAX(F4:H4)"
                    End If
                
                    If FirstString_WBeam_S = Keyword_WBeam_Btm Then
                        Exit Do
                    End If
                Loop
                
            End If
            
            '读取梁底筋
            FirstString_WBeam_S = Mid(data, 1, 7)
            If FirstString_WBeam_S = "% Steel" Then
                Debug.Print "读取" & CStr(num) & "层墙梁底筋配筋率……"
                Sheets("WBI").Cells(i_wb1 + 4, 7) = extractNumberFromString(data, 5)
            End If
            
            
            '读取梁配箍率
            If Mid(data, 2, 3) = Keyword_WBeam_Rsv Then
                Debug.Print "读取" & CStr(num) & "层墙梁配箍率……"
                Sheets("WBI").Cells(i_wb1 + 4, 12) = extractNumberFromString(data, 1)
                If Sheets("WBI").Cells(i_wb1 + 4, 12) < extractNumberFromString(data, 5) Then
                    Sheets("WBI").Cells(i_wb1 + 4, 12) = extractNumberFromString(data, 5)
                End If
                If Sheets("WBI").Cells(i_wb1 + 4, 12) < extractNumberFromString(data, 9) Then
                    Sheets("WBI").Cells(i_wb1 + 4, 12) = extractNumberFromString(data, 9)
                End If
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_wb1 = i_wb1 + 1
                Exit Do
            End If

        Loop
    End If

    
Loop

Close #1


'打开结果文件
Open (filepath2) For Input Access Read As #i


'===========================================================================================读取中震信息

Debug.Print "开始遍历中震结果文件wpj" & CStr(num); ".out; "
Debug.Print "读取相关指标"
Debug.Print "……"





Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
    
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Column = Mid(data, 2, 4)
    FirstString_Wall = Mid(data, 2, 5)
    FirstString_Beam = Mid(data, 2, 4)
    FirstString_WBeam = Mid(data, 2, 5)

    '--------------------------------------------------------------------------读取柱的信息
    If FirstString_Column = Keyword_Column Then
        Debug.Print "读取" & CStr(num) & "层柱信息……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Column_UC = Mid(data, 20, 3)
            If FirstString_Column_UC = Keyword_Column_UC Then
                '读取柱轴压比
                Sheets("CI").Cells(i_c2 + 4, 7) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                '读取柱主筋配筋率
                Sheets("CI").Cells(i_c2 + 4, 9) = Mid(data, 34, 6)
                '读取柱箍筋配筋率
                Sheets("CI").Cells(i_c2 + 4, 11) = Mid(data, 48, 6)
            End If
            
            If Mid(data, 2, 5) = Keyword_Column_V Then
                '读取抗剪承载力
                Sheets("CI").Cells(i_c2 + 4, 13) = extractNumberFromString(data, 1)
                Sheets("CI").Cells(i_c2 + 4, 15) = extractNumberFromString(data, 2)
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_c2 = i_c2 + 1
                Exit Do
            End If
            
        Loop
    End If
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        '读取墙截面
        B_w = Sheets("WCI").Cells(i_w2 + 4, 4)
        H_w = Sheets("WCI").Cells(i_w2 + 4, 5)
        Debug.Print "读取" & CStr(num) & "层墙信息……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Wall_UC = Mid(data, 20, 3)
            If FirstString_Wall_UC = Keyword_Wall_UC Then
                '读取墙轴压比
                Debug.Print "读取" & CStr(num) & "层墙轴压比……"
                Sheets("WCI").Cells(i_w2 + 4, 8) = StringfromStringforReg(data, "\s+0\.\d*", 1)
            End If
            
            If Mid(data, 7, 2) = "M=" And Mid(data, 19, 2) = "N=" Then
                '读取暗柱配筋面积
                 Debug.Print "读取" & CStr(num) & "层墙暗柱配筋……"
                Sheets("WCI").Cells(i_w2 + 4, 10) = extractNumberFromString(data, 4)
                As_az = Sheets("WCI").Cells(i_w2 + 4, 10)
                '暗柱配筋若为0，即构造配筋，改为1，方便小中震比值的计算
                If As_az = 0 Then
                    Sheets("WCI").Cells(i_w2 + 4, 10) = 1
                End If
            End If
            
            If Mid(data, 7, 2) = "V=" And Mid(data, 19, 2) = "N=" Then
                '读取水平分布筋配筋面积
                 Debug.Print "读取" & CStr(num) & "层墙水平配筋……"
                Sheets("WCI").Cells(i_w2 + 4, 12) = extractNumberFromString(data, 4)
                Sheets("WCI").Cells(i_w2 + 4, 14) = extractNumberFromString(data, 5)
            End If
            
            If Mid(data, 2, 5) = Keyword_Wall_V Then
                '读取抗剪承载力
                 Debug.Print "读取" & CStr(num) & "层墙抗剪承载力……"
                Sheets("WCI").Cells(i_w2 + 4, 16) = extractNumberFromString(data, 1)
                Sheets("WCI").Cells(i_w2 + 4, 18) = extractNumberFromString(data, 2)
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_w2 = i_w2 + 1
                Exit Do
            End If
        Loop
    End If
    
    '--------------------------------------------------------------------------读取梁信息
    If FirstString_Beam = Keyword_Beam Then
        Debug.Print "读取" & CStr(num) & "层梁信息……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_Beam_S = Mid(data, 1, 7)
            If FirstString_Beam_S = Keyword_Beam_Top Then
                '读取梁顶筋
                Do While Not EOF(1)
                    Line Input #1, data
                    FirstString_Beam_S = Mid(data, 1, 7)
                    If FirstString_Beam_S = "% Steel" Then
                        Debug.Print "读取" & CStr(num) & "层梁顶筋配筋率……"
                        Sheets("BI").Cells(i_b2 + 4, 9) = extractNumberFromString(data, 1)
                        Sheets("BI").Cells(i_b2 + 4, 11) = extractNumberFromString(data, 9)
'                        Sheets("BI").Cells(i_b1 + 4, 8) = "=MAX(F4:H4)"
                    End If
                
                    If FirstString_Beam_S = Keyword_Beam_Btm Then
                        Exit Do
                    End If
                Loop
                
            End If
            
            '读取梁底筋
            FirstString_Beam_S = Mid(data, 1, 7)
            If FirstString_Beam_S = "% Steel" Then
                Debug.Print "读取" & CStr(num) & "层梁底筋配筋率……"
                Sheets("BI").Cells(i_b2 + 4, 10) = extractNumberFromString(data, 5)
            End If
            
            
            '读取梁配箍率
            If Mid(data, 2, 3) = Keyword_Beam_Rsv Then
                Debug.Print "读取" & CStr(num) & "层梁配箍率……"
                Sheets("BI").Cells(i_b2 + 4, 13) = extractNumberFromString(data, 1)
                If Sheets("BI").Cells(i_b2 + 4, 13) < extractNumberFromString(data, 5) Then
                    Sheets("BI").Cells(i_b2 + 4, 13) = extractNumberFromString(data, 5)
                End If
                If Sheets("BI").Cells(i_b2 + 4, 13) < extractNumberFromString(data, 9) Then
                    Sheets("BI").Cells(i_b2 + 4, 13) = extractNumberFromString(data, 9)
                End If
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_b2 = i_b2 + 1
                Exit Do
            End If

        Loop
    End If
    
    
    '--------------------------------------------------------------------------读取墙梁信息
    If FirstString_WBeam = Keyword_WBeam Then
        Debug.Print "读取" & CStr(num) & "层墙梁信息……"
        Line Input #1, data
        Do While Not EOF(1)
            Line Input #1, data
            FirstString_WBeam_S = Mid(data, 1, 7)
            If FirstString_WBeam_S = Keyword_WBeam_Top Then
                '读取梁顶筋
                Do While Not EOF(1)
                    Line Input #1, data
                    FirstString_WBeam_S = Mid(data, 1, 7)
                    If FirstString_WBeam_S = "% Steel" Then
                        Debug.Print "读取" & CStr(num) & "层墙梁顶筋配筋率……"
                        Sheets("WBI").Cells(i_wb2 + 4, 9) = extractNumberFromString(data, 1)
                        Sheets("WBI").Cells(i_wb2 + 4, 11) = extractNumberFromString(data, 9)
'                        Sheets("WBI").Cells(i_wb2 + 4, 8) = "=MAX(F4:H4)"
                    End If
                
                    If FirstString_WBeam_S = Keyword_WBeam_Btm Then
                        Exit Do
                    End If
                Loop
                
            End If
            
            '读取梁底筋
            FirstString_WBeam_S = Mid(data, 1, 7)
            If FirstString_WBeam_S = "% Steel" Then
                Debug.Print "读取" & CStr(num) & "层墙梁底筋配筋率……"
                Sheets("WBI").Cells(i_wb2 + 4, 10) = extractNumberFromString(data, 5)
            End If
            
            
            '读取梁配箍率
            If Mid(data, 2, 3) = Keyword_WBeam_Rsv Then
                Debug.Print "读取" & CStr(num) & "层墙梁配箍率……"
                Sheets("WBI").Cells(i_wb2 + 4, 13) = extractNumberFromString(data, 1)
                If Sheets("WBI").Cells(i_wb2 + 4, 13) < extractNumberFromString(data, 5) Then
                    Sheets("WBI").Cells(i_wb2 + 4, 13) = extractNumberFromString(data, 5)
                End If
                If Sheets("WBI").Cells(i_wb2 + 4, 13) < extractNumberFromString(data, 9) Then
                    Sheets("WBI").Cells(i_wb2 + 4, 13) = extractNumberFromString(data, 9)
                End If
            End If
            
            If CheckRegExpfromString(data, "---") = True Then
                i_wb2 = i_wb2 + 1
                Exit Do
            End If

        Loop
    End If

    
Loop

Close #1




End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub AddInfoSheet()


'计算运行时间
Dim sngStart As Single
sngStart = Timer

Call Addsh("CI")
Call Addsh("WCI")
Call Addsh("BI")
Call Addsh("WBI")

'======================================================================================================添加表格Column的标题

'清除工作表所有内容
Sheets("CI").Cells.Clear


'加表格线
Call AddFormLine("CI", "A2:Q20000")

'加背景色
Call AddShadow("CI", "A2:Q3", 10092441)

With Sheets("CI")
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    .range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 6) = "柱配筋信息提取"
    .Cells(1, 6).Font.name = "黑体"
    .Cells(1, 6).Font.Size = "20"
    '合并单元格
    .range("F1:H1").MergeCells = True
    
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "柱编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 6) = "轴压比"
    .Cells(2, 8) = "主筋配筋率Rs"
    .Cells(2, 10) = "箍筋配筋率Rsv"
    .Cells(2, 12) = "CB-XF"
    .Cells(2, 14) = "CB-YF"
    .Cells(2, 16) = "中震/小震"
    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "小震"
    .Cells(3, 7) = "中震"
    .Cells(3, 8) = "小震"
    .Cells(3, 9) = "中震"
    .Cells(3, 10) = "小震"
    .Cells(3, 11) = "中震"
    .Cells(3, 12) = "小震"
    .Cells(3, 13) = "中震"
    .Cells(3, 14) = "小震"
    .Cells(3, 15) = "中震"
    .Cells(3, 16) = "主筋"
    .Cells(3, 17) = "箍筋"
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:E2").MergeCells = True
    .range("F2:G2").MergeCells = True
    .range("H2:I2").MergeCells = True
    .range("J2:K2").MergeCells = True
    .range("L2:M2").MergeCells = True
    .range("N2:O2").MergeCells = True
    .range("P2:Q2").MergeCells = True
End With

'冻结首行首列
Sheets("CI").Select
range("B4").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True



'======================================================================================================添加表格Wall的标题

'清除工作表所有内容
Sheets("WCI").Cells.Clear


'加表格线
Call AddFormLine("WCI", "A2:T20000")

'加背景色
'Call AddShadow("WCI", "A2:T3", 6750105)
Call AddShadow("WCI", "A2:T3", 10092441)


With Sheets("WCI")
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    .range("G4:T20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 9) = "墙配筋信息提取"
    .Cells(1, 9).Font.name = "黑体"
    .Cells(1, 9).Font.Size = "20"
    '合并单元格
    .range("I1:L1").MergeCells = True
    
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "柱编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 7) = "轴压比"
    .Cells(2, 9) = "一端暗柱配筋面积As"
    .Cells(2, 11) = "水平分布筋面积Ash"
    .Cells(2, 13) = "分布筋配筋率Rsh"
    .Cells(2, 15) = "CB-XF"
    .Cells(2, 17) = "CB-YF"
    .Cells(2, 19) = "中震/小震"
    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "Lwc"
    .Cells(3, 7) = "小震"
    .Cells(3, 8) = "中震"
    .Cells(3, 9) = "小震"
    .Cells(3, 10) = "中震"
    .Cells(3, 11) = "小震"
    .Cells(3, 12) = "中震"
    .Cells(3, 13) = "小震"
    .Cells(3, 14) = "中震"
    .Cells(3, 15) = "小震"
    .Cells(3, 16) = "中震"
    .Cells(3, 17) = "小震"
    .Cells(3, 18) = "中震"
    .Cells(3, 19) = "一端暗柱"
    .Cells(3, 20) = "水平分布"
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:F2").MergeCells = True
    .range("G2:H2").MergeCells = True
    .range("I2:J2").MergeCells = True
    .range("K2:L2").MergeCells = True
    .range("M2:N2").MergeCells = True
    .range("O2:P2").MergeCells = True
    .range("Q2:R2").MergeCells = True
    .range("S2:T2").MergeCells = True
End With

'冻结首行首列
Sheets("WCI").Select
range("B4").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

'======================================================================================================添加表格Beam的标题

'清除工作表所有内容
Sheets("BI").Cells.Clear


'加表格线
Call AddFormLine("BI", "A2:Q20000")

'加背景色
'Call AddShadow("BI", "A2:03", 6750105)
Call AddShadow("BI", "A2:Q3", 10092441)

With Sheets("BI")
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    .range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 6) = "梁配筋信息提取"
    .Cells(1, 6).Font.name = "黑体"
    .Cells(1, 6).Font.Size = "20"
    '合并单元格
    .range("F1:H1").MergeCells = True
    
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "柱编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 6) = "配筋率（小震）"
    .Cells(2, 9) = "配筋率（中震）"
    .Cells(2, 12) = "配箍率"
    .Cells(2, 14) = "中震/小震"
    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "I点顶筋"
    .Cells(3, 7) = "中点底筋"
    .Cells(3, 8) = "J点顶筋"
    .Cells(3, 9) = "I点顶筋"
    .Cells(3, 10) = "中点底筋"
    .Cells(3, 11) = "点顶筋"
    .Cells(3, 12) = "小震"
    .Cells(3, 13) = "中震"
    .Cells(3, 14) = "I点顶筋"
    .Cells(3, 15) = "中点底筋"
    .Cells(3, 16) = "J点顶筋"
    .Cells(3, 17) = "配箍"
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:E2").MergeCells = True
    .range("F2:H2").MergeCells = True
    .range("I2:K2").MergeCells = True
    .range("L2:M2").MergeCells = True
    .range("N2:Q2").MergeCells = True
    
End With

'冻结首行首列
Sheets("BI").Select
range("B4").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True



'======================================================================================================添加表格WBeam的标题

'清除工作表所有内容
Sheets("WBI").Cells.Clear


'加表格线
Call AddFormLine("WBI", "A2:Q20000")

'加背景色
'Call AddShadow("WBI", "A2:03", 6750105)
Call AddShadow("WBI", "A2:Q3", 10092441)

With Sheets("WBI")
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    .range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 6) = "墙梁配筋信息提取"
    .Cells(1, 6).Font.name = "黑体"
    .Cells(1, 6).Font.Size = "20"
    '合并单元格
    .range("F1:H1").MergeCells = True
    
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "柱编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 6) = "配筋率（小震）"
    .Cells(2, 9) = "配筋率（中震）"
    .Cells(2, 12) = "配箍率"
    .Cells(2, 14) = "中震/小震"
    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "I点顶筋"
    .Cells(3, 7) = "中点底筋"
    .Cells(3, 8) = "J点顶筋"
    .Cells(3, 9) = "I点顶筋"
    .Cells(3, 10) = "中点底筋"
    .Cells(3, 11) = "点顶筋"
    .Cells(3, 12) = "小震"
    .Cells(3, 13) = "中震"
    .Cells(3, 14) = "I点顶筋"
    .Cells(3, 15) = "中点底筋"
    .Cells(3, 16) = "J点顶筋"
    .Cells(3, 17) = "配箍"
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:E2").MergeCells = True
    .range("F2:H2").MergeCells = True
    .range("I2:K2").MergeCells = True
    .range("L2:M2").MergeCells = True
    .range("N2:Q2").MergeCells = True
    
End With



'冻结首行首列
Sheets("WBI").Select
range("B4").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

End Sub


