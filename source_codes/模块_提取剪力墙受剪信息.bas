Attribute VB_Name = "模块_提取剪力墙受剪信息"
Option Explicit
Public i_wa As Integer

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/7/9
'1.添加剪力墙部位选择，分别定义不同的配筋率。


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/4
'1.增加墙分布筋的配筋表格"wallshear"（暂时为不删除的隐藏的表格，配筋时索引查找）
'2.删除原来手工输入分布筋的语句
'3.增加PKPM按构件编号读取的模块
'4.增加抗剪截面要求的检查，构造配筋的检查
'5.抗剪截面要求的计算公式中混凝土强度采用设计值，除以gammaRE，即0.85
'6.增加读取所有墙编号总数的函数，分PKPM和YJK两个

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/3
'1.增加了YJK的模块
'2.分成按楼层和按编号两种读取方法
'3.原member_info2弃用，运行时在面板模块直接调用
'4.增加超限项读取

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/24
'更新内容:
'1.更正C50的ft值
'2.更正抗剪承载力公式

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/19
'更新内容:
'1.修正一些bug，i_wa变量只做写入序号用

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/11
'更新内容:
'1.将原来的member_info模块分成两个，member_info1和member_info2，前者生成表格、层号和构件编号，提供用户修改构件编号后，再运行member_info2进行数据读取
'2.添加beta_c变量
'3.添加钢筋排数变量

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/17
'更新内容:
'1.补充完整C30C80的材料强度值

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/9/11
'更新内容:
'1.

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                          提取剪力墙受剪信息代码                      ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


Sub member_info_f(path1 As String, mem As Integer, startf As Integer, endf As Integer, softname As String, infotype As String)


'-------------------------------------------------------------------------------------------- 建立工作表
Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Call Addsh(wallshearsheet)

'清除工作表所有内容
Sheets(wallshearsheet).Cells.Clear


'加表格线
Call AddFormLine(wallshearsheet, "A2:AC20000")

'加背景色
Call AddShadow(wallshearsheet, "A2:AC3", 10092441)

With Sheets(wallshearsheet)
   
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    '.range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    '调整单元格宽度
    .Columns("AB:AB").ColumnWidth = 18.13
    .Columns("X:Y").ColumnWidth = 11.88
    .Columns("W:W").ColumnWidth = 10.63
   
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 10) = "剪力墙抗剪验算"
    .Cells(1, 10).Font.name = "黑体"
    .Cells(1, 10).Font.Size = "20"
    '合并单元格
    .range("J1:M1").MergeCells = True
   
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "墙编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 7) = "轴压比"
    .Cells(2, 8) = "CB-XF"
    .Cells(2, 9) = "CB-YF"
    .Cells(2, 10) = "轴力"
    .Cells(2, 11) = "剪力"
    .Cells(2, 12) = "剪跨比"
    .Cells(2, 13) = "砼等级"
    .Cells(2, 14) = "aa"
    .Cells(2, 15) = "水平钢筋"
    .Cells(2, 18) = "材料强度"
    .Cells(2, 22) = "beta_c"
    .Cells(2, 23) = "修正后剪跨比"
    .Cells(2, 24) = "抗剪承载力"
    .Cells(2, 25) = "抗剪截面要求"
    .Cells(2, 26) = "抗剪截面要求检查"
    .Cells(2, 27) = "构造配筋检查"
    .Cells(2, 28) = "超限项检查"

    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "Lwc"

    .Cells(3, 8) = "kN"
    .Cells(3, 9) = "kN"
    .Cells(3, 10) = "kN"
    .Cells(3, 11) = "kN"

    .Cells(3, 15) = "直径"
    .Cells(3, 16) = "间距"
    .Cells(3, 17) = "排数"
    .Cells(3, 18) = "fc"
    .Cells(3, 19) = "ft"
    .Cells(3, 20) = "fck"
    .Cells(3, 21) = "fyv"
    .Cells(3, 24) = "kN"
    .Cells(3, 25) = "kN"


    
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:F2").MergeCells = True
    .range("G2:G3").MergeCells = True

    .range("L2:L3").MergeCells = True
    .range("M2:M3").MergeCells = True
    .range("n2:n3").MergeCells = True
    .range("o2:q2").MergeCells = True
    .range("r2:u2").MergeCells = True
    .range("v2:v3").MergeCells = True
    .range("w2:w3").MergeCells = True
    .range("z2:z3").MergeCells = True
    .range("aa2:aa3").MergeCells = True
    .range("ab2:ab3").MergeCells = True
    '.range("v2:v3").MergeCells = True
    '.range("w2:w3").MergeCells = True

    
End With

'冻结首行首列
Sheets(wallshearsheet).Select
range("d4").Select
With ActiveWindow
    .SplitColumn = 3
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

i_wa = 0

Dim i As Integer

For i = startf To endf

    Sheets(wallshearsheet).Cells(i + 3, 2) = i
    Sheets(wallshearsheet).Cells(i + 3, 3) = mem
'Call PKPM_Wall_Info(path1, mem, i)

Next

End Sub

'================================================================================================================
Sub member_info_m(path1 As String, flo As Integer, startm As Integer, endm As Integer, softname As String, infotype As String)


'-------------------------------------------------------------------------------------------- 建立工作表
Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Call Addsh(wallshearsheet)

'清除工作表所有内容
Sheets(wallshearsheet).Cells.Clear


'加表格线
Call AddFormLine(wallshearsheet, "A2:AF20000")

'加背景色
Call AddShadow(wallshearsheet, "A2:AF3", 10092441)

With Sheets(wallshearsheet)
   
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    '.range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    '调整单元格宽度
    .Columns("AB:AF").ColumnWidth = 18.13
    .Columns("X:Y").ColumnWidth = 11.88
    .Columns("W:W").ColumnWidth = 10.63
    
   
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 10) = "剪力墙抗剪验算"
    .Cells(1, 10).Font.name = "黑体"
    .Cells(1, 10).Font.Size = "20"
    '合并单元格
    .range("J1:M1").MergeCells = True
   
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 1) = "序号"
    .Cells(2, 2) = "楼层"
    .Cells(2, 3) = "墙编号"
    .Cells(2, 4) = "截面信息"
    .Cells(2, 7) = "轴压比"
    .Cells(2, 8) = "CB-XF"
    .Cells(2, 9) = "CB-YF"
    .Cells(2, 10) = "轴力"
    .Cells(2, 11) = "剪力"
    .Cells(2, 12) = "剪跨比"
    .Cells(2, 13) = "砼等级"
    .Cells(2, 14) = "aa"
    .Cells(2, 15) = "水平钢筋"
    .Cells(2, 18) = "材料强度"
    .Cells(2, 22) = "beta_c"
    .Cells(2, 23) = "修正后剪跨比"
    .Cells(2, 24) = "抗剪承载力"
    .Cells(2, 25) = "抗剪截面要求"
    .Cells(2, 26) = "抗剪截面要求检查"
    .Cells(2, 27) = "构造配筋检查"
    .Cells(2, 28) = "超限项检查"

    .Cells(3, 4) = "B"
    .Cells(3, 5) = "H"
    .Cells(3, 6) = "Lwc"

    .Cells(3, 8) = "kN"
    .Cells(3, 9) = "kN"
    .Cells(3, 10) = "kN"
    .Cells(3, 11) = "kN"

    .Cells(3, 15) = "直径"
    .Cells(3, 16) = "间距"
    .Cells(3, 17) = "排数"
    .Cells(3, 18) = "fc"
    .Cells(3, 19) = "ft"
    .Cells(3, 20) = "fck"
    .Cells(3, 21) = "fyv"
    .Cells(3, 24) = "kN"
    .Cells(3, 25) = "kN"


    
    '合并单元格
    .range("A2:A3").MergeCells = True
    .range("B2:B3").MergeCells = True
    .range("C2:C3").MergeCells = True
    .range("D2:F2").MergeCells = True
    .range("G2:G3").MergeCells = True

    .range("L2:L3").MergeCells = True
    .range("M2:M3").MergeCells = True
    .range("n2:n3").MergeCells = True
    .range("o2:q2").MergeCells = True
    .range("r2:u2").MergeCells = True
    .range("v2:v3").MergeCells = True
    .range("w2:w3").MergeCells = True
    .range("z2:z3").MergeCells = True
    .range("aa2:aa3").MergeCells = True
    .range("AB2:AF3").MergeCells = True
    '.range("v2:v3").MergeCells = True
    '.range("w2:w3").MergeCells = True

    
End With

'冻结首行首列
Sheets(wallshearsheet).Select
range("d4").Select
With ActiveWindow
    .SplitColumn = 3
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

i_wa = 0

Dim i As Integer

For i = startm To endm

    Sheets(wallshearsheet).Cells(i + 3, 2) = flo
    Sheets(wallshearsheet).Cells(i + 3, 3) = i
'Call PKPM_Wall_Info(path1, mem, i)

Next

End Sub

'====================================================================================================================================
'PKPM按楼层读取
Sub PKPM_Wall_Info_F(path1 As String, num As Integer, softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Dim n As Integer
n = num


'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String
'赋值
Keyword_Wall_UC = "Uc="


'柱、墙抗剪承载力行关键词
Dim Keyword_Wall_V As String
'赋值
Keyword_Wall_V = "抗剪承载力"


'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_Wall As String
'柱、墙轴压比，梁配筋率
Dim FirstString_Wall_UC As String


'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "WPJ" & CStr(num) & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename


Sheets(wallshearsheet).Select

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
    FirstString_Wall = Mid(data, 2, 5)
    
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        Debug.Print "读取" & CStr(num) & "层墙信息……"
        
        '读取墙编号
        mem = Sheets(wallshearsheet).Cells(n + 3, 3)
        
        If extractNumberFromString(data, 1) = mem Then
            '写入序号
            Sheets(wallshearsheet).Cells(n + 3, 1) = i_wa + 1
            '写入楼层号
            Sheets(wallshearsheet).Cells(n + 3, 2) = n
'            写入钢筋直径和间距
'            Sheets(wallshearsheet).Cells(n + 3, 15) = OUTReader_Main.D_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 16) = OUTReader_Main.DJ_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 17) = 2
        
            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallshearsheet).Cells(n + 3, 4) = extractNumberFromString(data, 4) * 1000
            B_w = Sheets(wallshearsheet).Cells(n + 3, 4)
            Sheets(wallshearsheet).Cells(n + 3, 5) = extractNumberFromString(data, 5) * 1000
            H_w = Sheets(wallshearsheet).Cells(n + 3, 5)
            Sheets(wallshearsheet).Cells(n + 3, 6) = extractNumberFromString(data, 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                FirstString_Wall_UC = Mid(data, 20, 3)
                If Mid(data, 2, 2) = "aa" Then
                    Sheets(wallshearsheet).Cells(n + 3, 14) = extractNumberFromString(data, 1)
                    Sheets(wallshearsheet).Cells(n + 3, 13) = extractNumberFromString(data, 3)
                    Sheets(wallshearsheet).Cells(n + 3, 21) = extractNumberFromString(data, 5)
                End If
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
                    Debug.Print "读取" & CStr(num) & "层墙轴压比……"
                    Sheets(wallshearsheet).Cells(n + 3, 7) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                End If
           
                If Mid(data, 7, 2) = "M=" And Mid(data, 19, 2) = "V=" Then
                    '读取暗柱配筋面积
                    Debug.Print "读取" & CStr(num) & "层剪跨比……"                   '----------------------------输出剪跨比
                    Sheets(wallshearsheet).Cells(n + 3, 12) = extractNumberFromString(data, 4)
                    Sheets(wallshearsheet).Cells(n + 3, 12) = Round(Sheets(wallshearsheet).Cells(n + 3, 12), 3)
                End If
           

           
                If Mid(data, 7, 2) = "V=" And Mid(data, 19, 2) = "N=" Then
                    '读取水平分布筋配筋面积
                    Debug.Print "读取" & CStr(num) & "层墙内力……"
                    Sheets(wallshearsheet).Cells(n + 3, 11) = extractNumberFromString(data, 2) '----------------------------输出剪力
                    Sheets(wallshearsheet).Cells(n + 3, 11) = Round(Abs(Sheets(wallshearsheet).Cells(n + 3, 11)), 0)
                    Sheets(wallshearsheet).Cells(n + 3, 10) = extractNumberFromString(data, 3) '----------------------------输出轴力
                    Sheets(wallshearsheet).Cells(n + 3, 10) = Round(Sheets(wallshearsheet).Cells(n + 3, 10), 0)
                    If Sheets(wallshearsheet).Cells(n + 3, 10) > 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 10).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 2).Interior.ColorIndex = 3
                    End If
                End If
           
                If Mid(data, 2, 5) = Keyword_Wall_V Then
                    '读取抗剪承载力
                     Debug.Print "读取" & CStr(num) & "层墙抗剪承载力……"
                    Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                    Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                End If

                If Sheets(wallshearsheet).Cells(n + 3, 13) = "80" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 35.9
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.22
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 50.2
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.8
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "75" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 33.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.18
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 47.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.83
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "70" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 31.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.14
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 44.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.87
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "65" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 29.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.09
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 41.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.9
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "60" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 27.5
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.04
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 38.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.93
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "55" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 25.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.96
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 35.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.97
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "50" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 23.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.89
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 32.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "45" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 21.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.8
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 29.6
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "40" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 19.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.71
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 26.8
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "35" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 16.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.57
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 23.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "30" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 14.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.43
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 20.1
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                Else
                    MsgBox ("请自行输入强度值！")
                End If
                
                
                If Sheets(wallshearsheet).Cells(n + 3, 12) < 1.5 Then
                   Sheets(wallshearsheet).Cells(n + 3, 23) = 1.5
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 12) > 2.2 Then
                    Sheets(wallshearsheet).Cells(n + 3, 23) = 2.2
                Else
                    Sheets(wallshearsheet).Cells(n + 3, 23) = Sheets(wallshearsheet).Cells(n + 3, 12)
                End If
                
                With Sheets(wallshearsheet)
                
                '抗剪截面要求检查、构造配筋检查
                .Cells(n + 3, 26).Formula = "=RC[-15]/RC[-1]"
                .Cells(n + 3, 27).Formula = "=RC[-16]/RC[-3]"
                
                '分布筋按构造配筋
                If OUTReader_Main.WallLocation1 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C3:R26C3)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C4:R26C4)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C5:R26C5)"
                ElseIf OUTReader_Main.WallLocation2 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C15:R26C15)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C16:R26C16)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C17:R26C17)"
                End If
                
                '抗剪承载力计算、抗剪截面要求计算
                If .Cells(n + 3, 10) < 0 Then
                    
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])+0.1*MIN(ABS(RC[-14]*1000),0.2*RC[-6]*RC[-20]*RC[-19]))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                
                Else
                
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])-0.1*ABS(RC[-14]*1000))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                End If
                
                .Cells(n + 3, 25).Formula = "=0.15 *RC[-3]*RC[-7] * RC[-21] * (RC[-20] - RC[-11]) / 1000 / 0.85"
                End With
                
                Sheets(wallshearsheet).Columns("V:V").NumberFormatLocal = "0.00"
                Sheets(wallshearsheet).Columns("W:AA").NumberFormatLocal = "0.00"
                        
                If CheckRegExpfromString(data, "---") = True Then
                    i_wa = i_wa + 1
                    Exit Do
                End If
            Loop
        End If
        
        '抗剪截面要求和构造配筋检查不满足时标黄色
        With Sheets(wallshearsheet)
            If .Cells(n + 3, 26) >= 1 Then
                .Cells(n + 3, 26).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If

            If .Cells(n + 3, 27) >= 1 Then
                .Cells(n + 3, 27).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If
        End With
        
    End If
   
   
Loop

Close #1

End Sub

'==============================================================================================================================================================================

'PKPM按楼层读取
Sub PKPM_Wall_Info_M(path1 As String, num As Integer, softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Dim n As Integer
n = num


'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

Dim flo As Integer


'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String
'赋值
Keyword_Wall_UC = "Uc="


'柱、墙抗剪承载力行关键词
Dim Keyword_Wall_V As String
'赋值
Keyword_Wall_V = "抗剪承载力"


'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_Wall As String
'柱、墙轴压比，梁配筋率
Dim FirstString_Wall_UC As String


'==========================================================================================生成文件读取路径

flo = Sheets(wallshearsheet).Cells(n + 3, 2)

'指定文件名为wpj_Num.out
Filename = "WPJ" & flo & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename


Sheets(wallshearsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历小震结果文件wpj" & flo & ".out; "
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Wall = Mid(data, 2, 5)
    
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        Debug.Print "读取" & flo & "层墙信息……"
        
        '读取墙编号
        mem = Sheets(wallshearsheet).Cells(n + 3, 3)
        
        If extractNumberFromString(data, 1) = mem Then
            '写入序号
            Sheets(wallshearsheet).Cells(n + 3, 1) = i_wa + 1

'            写入钢筋直径和间距
'            Sheets(wallshearsheet).Cells(n + 3, 15) = OUTReader_Main.D_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 16) = OUTReader_Main.DJ_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 17) = 2
        
            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallshearsheet).Cells(n + 3, 4) = extractNumberFromString(data, 4) * 1000
            B_w = Sheets(wallshearsheet).Cells(n + 3, 4)
            Sheets(wallshearsheet).Cells(n + 3, 5) = extractNumberFromString(data, 5) * 1000
            H_w = Sheets(wallshearsheet).Cells(n + 3, 5)
            Sheets(wallshearsheet).Cells(n + 3, 6) = extractNumberFromString(data, 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                FirstString_Wall_UC = Mid(data, 20, 3)
                If Mid(data, 2, 2) = "aa" Then
                    Sheets(wallshearsheet).Cells(n + 3, 14) = extractNumberFromString(data, 1)
                    Sheets(wallshearsheet).Cells(n + 3, 13) = extractNumberFromString(data, 3)
                    Sheets(wallshearsheet).Cells(n + 3, 21) = extractNumberFromString(data, 5)
                End If
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
                    Debug.Print "读取" & mem & "号墙轴压比……"
                    Sheets(wallshearsheet).Cells(n + 3, 7) = StringfromStringforReg(data, "\s+0\.\d*", 1)
                End If
           
                If Mid(data, 7, 2) = "M=" And Mid(data, 19, 2) = "V=" Then
                    '读取暗柱配筋面积
                    Debug.Print "读取" & mem & "号剪跨比……"                   '----------------------------输出剪跨比
                    Sheets(wallshearsheet).Cells(n + 3, 12) = extractNumberFromString(data, 4)
                    Sheets(wallshearsheet).Cells(n + 3, 12) = Round(Sheets(wallshearsheet).Cells(n + 3, 12), 3)
                End If
           

           
                If Mid(data, 7, 2) = "V=" And Mid(data, 19, 2) = "N=" Then
                    '读取水平分布筋配筋面积
                    Debug.Print "读取" & mem & "号墙内力……"
                    Sheets(wallshearsheet).Cells(n + 3, 11) = extractNumberFromString(data, 2) '----------------------------输出剪力
                    Sheets(wallshearsheet).Cells(n + 3, 11) = Round(Abs(Sheets(wallshearsheet).Cells(n + 3, 11)), 0)
                    Sheets(wallshearsheet).Cells(n + 3, 10) = extractNumberFromString(data, 3) '----------------------------输出轴力
                    Sheets(wallshearsheet).Cells(n + 3, 10) = Round(Sheets(wallshearsheet).Cells(n + 3, 10), 0)
                    If Sheets(wallshearsheet).Cells(n + 3, 10) > 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 10).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 2).Interior.ColorIndex = 3
                    End If
                End If
           
                If Mid(data, 2, 5) = Keyword_Wall_V Then
                    '读取抗剪承载力
                     Debug.Print "读取" & mem & "号墙抗剪承载力……"
                    Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                    Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                End If

                If Sheets(wallshearsheet).Cells(n + 3, 13) = "80" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 35.9
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.22
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 50.2
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.8
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "75" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 33.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.18
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 47.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.83
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "70" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 31.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.14
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 44.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.87
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "65" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 29.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.09
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 41.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.9
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "60" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 27.5
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.04
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 38.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.93
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "55" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 25.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.96
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 35.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.97
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "50" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 23.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.89
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 32.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "45" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 21.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.8
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 29.6
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "40" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 19.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.71
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 26.8
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "35" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 16.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.57
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 23.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "30" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 14.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.43
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 20.1
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                Else
                    MsgBox ("请自行输入强度值！")
                End If
                
                
                If Sheets(wallshearsheet).Cells(n + 3, 12) < 1.5 Then
                   Sheets(wallshearsheet).Cells(n + 3, 23) = 1.5
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 12) > 2.2 Then
                    Sheets(wallshearsheet).Cells(n + 3, 23) = 2.2
                Else
                    Sheets(wallshearsheet).Cells(n + 3, 23) = Sheets(wallshearsheet).Cells(n + 3, 12)
                End If
                
                With Sheets(wallshearsheet)
                
                '抗剪截面要求检查、构造配筋检查
                .Cells(n + 3, 26).Formula = "=RC[-15]/RC[-1]"
                .Cells(n + 3, 27).Formula = "=RC[-16]/RC[-3]"
                
                '分布筋按构造配筋
                If OUTReader_Main.WallLocation1 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C3:R26C3)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C4:R26C4)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C5:R26C5)"
                ElseIf OUTReader_Main.WallLocation2 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C15:R26C15)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C16:R26C16)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C17:R26C17)"
                End If
                
                '抗剪承载力计算、抗剪截面要求计算
                If .Cells(n + 3, 10) < 0 Then
                    
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])+0.1*MIN(ABS(RC[-14]*1000),0.2*RC[-6]*RC[-20]*RC[-19]))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                
                Else
                
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])-0.1*ABS(RC[-14]*1000))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                End If
                
                .Cells(n + 3, 25).Formula = "=0.15 *RC[-3]*RC[-7] * RC[-21] * (RC[-20] - RC[-11]) / 1000 / 0.85"
                End With
                
                Sheets(wallshearsheet).Columns("V:V").NumberFormatLocal = "0.00"
                Sheets(wallshearsheet).Columns("W:AA").NumberFormatLocal = "0.00"
                        
                If CheckRegExpfromString(data, "---") = True Then
                    i_wa = i_wa + 1
                    Exit Do
                End If
            Loop
        End If
        
        '抗剪截面要求和构造配筋检查不满足时标黄色
        With Sheets(wallshearsheet)
            If .Cells(n + 3, 26) >= 1 Then
                .Cells(n + 3, 26).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If

            If .Cells(n + 3, 27) >= 1 Then
                .Cells(n + 3, 27).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If
        End With
        
    End If
   
   
Loop

Close #1

End Sub


'==============================================================================================================================================================================
Sub YJK_Wall_Info_F(path1 As String, num As Integer, softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Dim n As Integer
n = num


'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String
'赋值
Keyword_Wall_UC = "Uc="


'柱、墙抗剪承载力行关键词
Dim Keyword_Wall_V As String
'赋值
Keyword_Wall_V = "抗剪承载力"


'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_Wall As String
'柱、墙轴压比，梁配筋率
Dim FirstString_Wall_UC As String


'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wpj" & CStr(num) & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename


Sheets(wallshearsheet).Select

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
    FirstString_Wall = Mid(data, 3, 5)
    
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        Debug.Print "读取" & CStr(num) & "层墙信息……"
        
        '读取墙编号
        mem = Sheets(wallshearsheet).Cells(n + 3, 3)
'        Debug.Print data
'        Debug.Print StringfromStringforReg(data, "\d+", 1)
        If StringfromStringforReg(data, "\d+", 1) = mem Then
            Debug.Print "写入序号"
            '写入序号
            Sheets(wallshearsheet).Cells(n + 3, 1) = i_wa + 1
            '写入楼层号
            Sheets(wallshearsheet).Cells(n + 3, 2) = n
'            写入钢筋直径和间距
'            Sheets(wallshearsheet).Cells(n + 3, 15) = OUTReader_Main.D_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 16) = OUTReader_Main.DJ_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 17) = 2
        
            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallshearsheet).Cells(n + 3, 4) = StringfromStringforReg(data, "\d+\.?\d*", 4) * 1000
            B_w = Sheets(wallshearsheet).Cells(n + 3, 4)
            Sheets(wallshearsheet).Cells(n + 3, 5) = StringfromStringforReg(data, "\d+\.?\d*", 5) * 1000
            H_w = Sheets(wallshearsheet).Cells(n + 3, 5)
            Sheets(wallshearsheet).Cells(n + 3, 6) = StringfromStringforReg(data, "\d+\.?\d*", 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                If Mid(data, 3, 5) = "Cover" Then
                    Sheets(wallshearsheet).Cells(n + 3, 14) = StringfromStringforReg(data, "\d+\.?\d*", 2)
                    Sheets(wallshearsheet).Cells(n + 3, 13) = StringfromStringforReg(data, "\d+\.?\d*", 5)
                    Sheets(wallshearsheet).Cells(n + 3, 21) = StringfromStringforReg(data, "\d+\.?\d*", 7)
                End If
                FirstString_Wall_UC = Mid(data, 22, 3)
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
                    Debug.Print "读取" & CStr(num) & "层墙轴压比……"
                    Sheets(wallshearsheet).Cells(n + 3, 7) = StringfromStringforReg(data, "0\.\d*", 1)
                End If
           
                If Mid(data, 8, 2) = "M=" And Mid(data, 21, 2) = "V=" Then
                    '读取暗柱配筋面积
                    Debug.Print "读取" & CStr(num) & "层剪跨比……"                   '----------------------------输出剪跨比
                    Sheets(wallshearsheet).Cells(n + 3, 12) = extractNumberFromString(data, 4)
                    Sheets(wallshearsheet).Cells(n + 3, 12) = Round(Sheets(wallshearsheet).Cells(n + 3, 12), 3)
                End If
           
           
                If Mid(data, 8, 2) = "V=" And Mid(data, 21, 2) = "N=" Then
                    '读取水平分布筋配筋面积
                    Debug.Print "读取" & CStr(num) & "层墙内力……"
                    Sheets(wallshearsheet).Cells(n + 3, 11) = extractNumberFromString(data, 2) '----------------------------输出剪力
                    Sheets(wallshearsheet).Cells(n + 3, 11) = Round(Abs(Sheets(wallshearsheet).Cells(n + 3, 11)), 0)
                    Sheets(wallshearsheet).Cells(n + 3, 10) = extractNumberFromString(data, 3) '----------------------------输出轴力
                    Sheets(wallshearsheet).Cells(n + 3, 10) = Round(Sheets(wallshearsheet).Cells(n + 3, 10), 0)
                    If Sheets(wallshearsheet).Cells(n + 3, 10) > 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 10).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 2).Interior.ColorIndex = 3
                    End If
                End If
                
                '检查超限项
                If Mid(data, 3, 2) = "**" Then
                
                    Sheets(wallshearsheet).Cells(n + 3, 28 + j) = data
                    If Sheets(wallshearsheet).Cells(n + 3, 28 + j) <> 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 28 + j).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 2).Interior.ColorIndex = 3
                        Else
                            Sheets(wallshearsheet).Cells(n + 3, 28 + j) = "无"
                    End If
                    '检查超限项，循环读取多个超限项
                    j = 1
                    Do While Not EOF(1)
                        Line Input #1, data
                        If Mid(data, 3, 2) = "**" Then
                            Sheets(wallshearsheet).Cells(n + 3, 28 + j) = data
                            If Sheets(wallshearsheet).Cells(n + 3, 28 + j) <> 0 Then
                                Sheets(wallshearsheet).Cells(n + 3, 28 + j).Interior.ColorIndex = 3
                                Sheets(wallshearsheet).Cells(n + 3, 2).Interior.ColorIndex = 3
                                Else
                                Sheets(wallshearsheet).Cells(n + 3, 28 + j) = "无"
                            End If
                            j = j + 1
                        End If
                        
                    '读取抗剪承载力，有超限项时，作为退出循环的判断依据
                        If Mid(data, 3, 5) = Keyword_Wall_V Then
                            Debug.Print "读取" & CStr(num) & "号墙抗剪承载力……"
                            Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                            Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                            Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                            Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                            Exit Do
                        End If
                    Loop
                    
                End If
                  
                '读取抗剪承载力
                If Mid(data, 3, 5) = Keyword_Wall_V Then
                    Debug.Print "读取" & CStr(num) & "号墙抗剪承载力……"
                    Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                    Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                    Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                        
                End If

                If Sheets(wallshearsheet).Cells(n + 3, 13) = "80" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 35.9
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.22
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 50.2
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.8
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "75" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 33.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.18
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 47.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.83
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "70" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 31.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.14
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 44.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.87
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "65" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 29.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.09
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 41.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.9
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "60" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 27.5
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.04
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 38.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.93
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "55" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 25.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.96
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 35.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.97
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "50" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 23.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.89
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 32.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "45" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 21.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.8
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 29.6
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "40" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 19.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.71
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 26.8
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "35" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 16.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.57
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 23.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "30" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 14.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.43
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 20.1
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                Else
                    MsgBox ("请自行输入强度值！")
                End If
                
                
                If Sheets(wallshearsheet).Cells(n + 3, 12) < 1.5 Then
                   Sheets(wallshearsheet).Cells(n + 3, 23) = 1.5
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 12) > 2.2 Then
                    Sheets(wallshearsheet).Cells(n + 3, 23) = 2.2
                Else
                    Sheets(wallshearsheet).Cells(n + 3, 23) = Sheets(wallshearsheet).Cells(n + 3, 12)
                End If
                
                With Sheets(wallshearsheet)
                
                '抗剪截面要求检查、构造配筋检查
                .Cells(n + 3, 26).Formula = "=RC[-15]/RC[-1]"
                .Cells(n + 3, 27).Formula = "=RC[-16]/RC[-3]"
                
                '分布筋按构造配筋
                If OUTReader_Main.WallLocation1 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C3:R26C3)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C4:R26C4)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C5:R26C5)"
                ElseIf OUTReader_Main.WallLocation2 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C15:R26C15)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C16:R26C16)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C17:R26C17)"
                End If
                
                '抗剪承载力计算、抗剪截面要求计算
                If .Cells(n + 3, 10) < 0 Then
                    
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])+0.1*MIN(ABS(RC[-14]*1000),0.2*RC[-6]*RC[-20]*RC[-19]))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                
                Else
                
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])-0.1*ABS(RC[-14]*1000))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                End If
                
                .Cells(n + 3, 25).Formula = "=0.15 *RC[-3]*RC[-7] * RC[-21] * (RC[-20] - RC[-11]) / 1000 / 0.85"
                End With
                
                Sheets(wallshearsheet).Columns("V:V").NumberFormatLocal = "0.00"
                Sheets(wallshearsheet).Columns("W:AA").NumberFormatLocal = "0.00"
                        
                If CheckRegExpfromString(data, "---") = True Then
                    i_wa = i_wa + 1
                    Exit Do
                End If
            Loop
        End If
        
        '抗剪截面要求和构造配筋检查不满足时标黄色
        With Sheets(wallshearsheet)
            If .Cells(n + 3, 26) >= 1 Then
                .Cells(n + 3, 26).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If

            If .Cells(n + 3, 27) >= 1 Then
                .Cells(n + 3, 27).Interior.ColorIndex = 6
                .Cells(n + 3, 2).Interior.ColorIndex = 3
            End If
        End With
        
    End If
   
   
Loop

Close #1

End Sub


'==============================================================================================================================================================================
Sub YJK_Wall_Info_M(path1 As String, num As Integer, softname As String, infotype As String)

Dim wallshearsheet As String
wallshearsheet = "WS_" & softname & "_" & infotype

Dim n As Integer
n = num

'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义楼层变量
Dim flo As Integer

'定义构件编号变量
Dim mem As Integer

'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String
'赋值
Keyword_Wall_UC = "Uc="


'柱、墙抗剪承载力行关键词
Dim Keyword_Wall_V As String
'赋值
Keyword_Wall_V = "抗剪承载力"


'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_Wall As String
'柱、墙轴压比，梁配筋率
Dim FirstString_Wall_UC As String


'==========================================================================================生成文件读取路径

flo = Sheets(wallshearsheet).Cells(n + 3, 2)

'指定文件名为wpj_Num.out
Filename = "wpj" & flo & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename


Sheets(wallshearsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历小震结果文件wpj" & flo & ".out; "
Debug.Print "读取相关指标"
Debug.Print "……"

'读取墙编号
mem = Sheets(wallshearsheet).Cells(n + 3, 3)

If mem <> 0 Then


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱、墙
    FirstString_Wall = Mid(data, 3, 5)
    
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
        Debug.Print "读取" & flo & "号墙信息……"
        
'        Debug.Print data
'        Debug.Print StringfromStringforReg(data, "\d+", 1)

        If StringfromStringforReg(data, "\d+", 1) = mem Then
            Debug.Print "写入序号"
            '写入序号
            Sheets(wallshearsheet).Cells(n + 3, 1) = i_wa + 1
'            写入钢筋直径和间距
'            Sheets(wallshearsheet).Cells(n + 3, 15) = OUTReader_Main.D_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 16) = OUTReader_Main.DJ_TextBox.Text
'            Sheets(wallshearsheet).Cells(n + 3, 17) = 2
        
            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallshearsheet).Cells(n + 3, 4) = StringfromStringforReg(data, "\d+\.?\d*", 4) * 1000
            B_w = Sheets(wallshearsheet).Cells(n + 3, 4)
            Sheets(wallshearsheet).Cells(n + 3, 5) = StringfromStringforReg(data, "\d+\.?\d*", 5) * 1000
            H_w = Sheets(wallshearsheet).Cells(n + 3, 5)
            Sheets(wallshearsheet).Cells(n + 3, 6) = StringfromStringforReg(data, "\d+\.?\d*", 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                If Mid(data, 3, 5) = "Cover" Then
                    Sheets(wallshearsheet).Cells(n + 3, 14) = StringfromStringforReg(data, "\d+\.?\d*", 2)
                    Sheets(wallshearsheet).Cells(n + 3, 13) = StringfromStringforReg(data, "\d+\.?\d*", 5)
                    Sheets(wallshearsheet).Cells(n + 3, 21) = StringfromStringforReg(data, "\d+\.?\d*", 7)
                End If
                FirstString_Wall_UC = Mid(data, 22, 3)
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
                    Debug.Print "读取" & CStr(num) & "号墙轴压比……"
                    Sheets(wallshearsheet).Cells(n + 3, 7) = StringfromStringforReg(data, "0\.\d*", 1)
                End If
           
                If Mid(data, 8, 2) = "M=" And Mid(data, 21, 2) = "V=" Then
                    '读取暗柱配筋面积
                    Debug.Print "读取" & CStr(num) & "号剪跨比……"                   '----------------------------输出剪跨比
                    Sheets(wallshearsheet).Cells(n + 3, 12) = extractNumberFromString(data, 4)
                    Sheets(wallshearsheet).Cells(n + 3, 12) = Round(Sheets(wallshearsheet).Cells(n + 3, 12), 3)
                End If
           
           
                If Mid(data, 8, 2) = "V=" And Mid(data, 21, 2) = "N=" Then
                    '读取水平分布筋配筋面积
                    Debug.Print "读取" & CStr(num) & "号墙内力……"
                    Sheets(wallshearsheet).Cells(n + 3, 11) = extractNumberFromString(data, 2) '----------------------------输出剪力
                    Sheets(wallshearsheet).Cells(n + 3, 11) = Round(Abs(Sheets(wallshearsheet).Cells(n + 3, 11)), 0)
                    Sheets(wallshearsheet).Cells(n + 3, 10) = extractNumberFromString(data, 3) '----------------------------输出轴力
                    Sheets(wallshearsheet).Cells(n + 3, 10) = Round(Sheets(wallshearsheet).Cells(n + 3, 10), 0)
                    If Sheets(wallshearsheet).Cells(n + 3, 10) > 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 10).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 3).Interior.ColorIndex = 3
                    End If
                End If
           

                If Mid(data, 3, 2) = "**" Then
                
                    Sheets(wallshearsheet).Cells(n + 3, 28 + j) = data
                    If Sheets(wallshearsheet).Cells(n + 3, 28 + j) <> 0 Then
                        Sheets(wallshearsheet).Cells(n + 3, 28 + j).Interior.ColorIndex = 3
                        Sheets(wallshearsheet).Cells(n + 3, 3).Interior.ColorIndex = 3
                        Else
                            Sheets(wallshearsheet).Cells(n + 3, 28 + j) = "无"
                    End If
                    '检查超限项，循环读取多个超限项
                    j = 1
                    Do While Not EOF(1)
                        Line Input #1, data
                        If Mid(data, 3, 2) = "**" Then
                            Sheets(wallshearsheet).Cells(n + 3, 28 + j) = data
                            If Sheets(wallshearsheet).Cells(n + 3, 28 + j) <> 0 Then
                                Sheets(wallshearsheet).Cells(n + 3, 28 + j).Interior.ColorIndex = 3
                                Sheets(wallshearsheet).Cells(n + 3, 3).Interior.ColorIndex = 3
                                Else
                                Sheets(wallshearsheet).Cells(n + 3, 28 + j) = "无"
                            End If
                            j = j + 1
                        End If
                        
                    '读取抗剪承载力
                        If Mid(data, 3, 5) = Keyword_Wall_V Then
                            Debug.Print "读取" & CStr(num) & "号墙抗剪承载力……"
                            Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                            Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                            Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                            Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                            Exit Do
                        End If
                    Loop
                    
                End If
                  
                '读取抗剪承载力
                    If Mid(data, 3, 5) = Keyword_Wall_V Then
                        Debug.Print "读取" & CStr(num) & "号墙抗剪承载力……"
                        Sheets(wallshearsheet).Cells(n + 3, 8) = extractNumberFromString(data, 1)
                        Sheets(wallshearsheet).Cells(n + 3, 8) = Round(Sheets(wallshearsheet).Cells(n + 3, 8), 1)
                        Sheets(wallshearsheet).Cells(n + 3, 9) = extractNumberFromString(data, 2)
                        Sheets(wallshearsheet).Cells(n + 3, 9) = Round(Sheets(wallshearsheet).Cells(n + 3, 9), 1)
                        
                    End If
                    
                If Sheets(wallshearsheet).Cells(n + 3, 13) = "80" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 35.9
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.22
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 50.2
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.8
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "75" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 33.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.18
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 47.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.83
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "70" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 31.8
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.14
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 44.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.87
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "65" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 29.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.09
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 41.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.9
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "60" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 27.5
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 2.04
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 38.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.93
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "55" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 25.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.96
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 35.5
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 0.97
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "50" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 23.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.89
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 32.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "45" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 21.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.8
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 29.6
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "40" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 19.1
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.71
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 26.8
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "35" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 16.7
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.57
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 23.4
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 13) = "30" Then
                   Sheets(wallshearsheet).Cells(n + 3, 18) = 14.3
                   Sheets(wallshearsheet).Cells(n + 3, 19) = 1.43
                   Sheets(wallshearsheet).Cells(n + 3, 20) = 20.1
                   Sheets(wallshearsheet).Cells(n + 3, 22) = 1#
                Else
                    MsgBox ("请自行输入强度值！")
                End If
                
                
                If Sheets(wallshearsheet).Cells(n + 3, 12) < 1.5 Then
                   Sheets(wallshearsheet).Cells(n + 3, 23) = 1.5
                ElseIf Sheets(wallshearsheet).Cells(n + 3, 12) > 2.2 Then
                    Sheets(wallshearsheet).Cells(n + 3, 23) = 2.2
                Else
                    Sheets(wallshearsheet).Cells(n + 3, 23) = Sheets(wallshearsheet).Cells(n + 3, 12)
                End If
                
                With Sheets(wallshearsheet)
                
                '抗剪截面要求检查、构造配筋检查
                .Cells(n + 3, 26).Formula = "=RC[-15]/RC[-1]"
                .Cells(n + 3, 27).Formula = "=RC[-16]/RC[-3]"
                
                '分布筋按构造配筋
                '分布筋按构造配筋
                If OUTReader_Main.WallLocation1 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C3:R26C3)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C4:R26C4)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C2:R26C2,wallrebar!R4C5:R26C5)"
                ElseIf OUTReader_Main.WallLocation2 Then
                    .Cells(n + 3, 15).Formula = "=LOOKUP(ROUNDUP(RC[-11]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C15:R26C15)"
                    .Cells(n + 3, 16).Formula = "=LOOKUP(ROUNDUP(RC[-12]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C16:R26C16)"
                    .Cells(n + 3, 17).Formula = "=LOOKUP(ROUNDUP(RC[-13]/50,0)*50,wallrebar!R4C14:R26C14,wallrebar!R4C17:R26C17)"
                End If
                
                '抗剪承载力计算、抗剪截面要求计算
                If .Cells(n + 3, 10) < 0 Then
                    
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])+0.1*MIN(ABS(RC[-14]*1000),0.2*RC[-6]*RC[-20]*RC[-19]))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                
                Else
                
                .Cells(n + 3, 24).Formula = "=1/0.85*(1/(RC[-1]-0.5)*(0.4*RC[-5]*RC[-20]*(RC[-19]-RC[-10])-0.1*ABS(RC[-14]*1000))+RC[-7]*0.8*RC[-3]*0.7854*RC[-9]^2/RC[-8]*(RC[-19]-RC[-10]))/1000"
                
                End If
                
                .Cells(n + 3, 25).Formula = "=0.15 *RC[-3]*RC[-7] * RC[-21] * (RC[-20] - RC[-11]) / 1000 / 0.85"
                End With
                
                                Sheets(wallshearsheet).Columns("V:V").NumberFormatLocal = "0.00"
                Sheets(wallshearsheet).Columns("V:V").NumberFormatLocal = "0.00"
                Sheets(wallshearsheet).Columns("W:AA").NumberFormatLocal = "0.00"
                        
                If CheckRegExpfromString(data, "---") = True Then
                    i_wa = i_wa + 1
                    Exit Do
                End If
            Loop
        End If
        
        '抗剪截面要求和构造配筋检查不满足时标黄色
        With Sheets(wallshearsheet)
            If .Cells(n + 3, 26) >= 1 Then
                .Cells(n + 3, 26).Interior.ColorIndex = 6
                .Cells(n + 3, 3).Interior.ColorIndex = 3
            End If

            If .Cells(n + 3, 27) >= 1 Then
                .Cells(n + 3, 27).Interior.ColorIndex = 6
                .Cells(n + 3, 3).Interior.ColorIndex = 3
            End If
        End With
        
    End If
   
   
Loop


End If

Close #1

End Sub



Function allmem_Y(path1 As String, num_floor As Integer) As Integer


allmem_Y = 0

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer
'指定文件名为wpj_Num.out
Filename = "wpj" & num_floor & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wpj" & num_floor & ".out; "
Debug.Print "……"

Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring
'    Debug.Print data
    '--------------------------------------------------------------------------读取墙的信息
    If Mid(data, 3, 5) = "N-WC=" Then
'        Debug.Print "test"
        allmem_Y = allmem_Y + 1
    End If

Loop

Close #1
    
End Function

Function allmem_P(path1 As String, num_floor As Integer) As Integer


allmem_P = 0

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer
'指定文件名为wpj_Num.out
Filename = "WPJ" & num_floor & ".OUT"

'生成完整文件路径
filepath1 = path1 & "\" & Filename

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wpj" & num_floor & ".out; "
Debug.Print "……"

Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring
'    Debug.Print data
    '--------------------------------------------------------------------------读取墙的信息
    If Mid(data, 2, 5) = "N-WC=" Then
'        Debug.Print "test"
        allmem_P = allmem_P + 1
    End If

Loop

Close #1
    
End Function



Sub wallrebar()

Dim i As Integer

Dim shname As String
shname = "wallrebar_test"

Call Addsh(shname)

'清除工作表所有内容
Sheets(shname).Cells.Clear


'加表格线
Call AddFormLine(shname, "B2:J26")

'加背景色
Call AddShadow(shname, "B2:J3", 10092441)

With Sheets(shname)
   
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    '.range("F4:Q20000").NumberFormatLocal = "0.00"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
   
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 5) = "剪力墙分布筋构造配筋"
    .Cells(1, 5).Font.name = "黑体"
    .Cells(1, 5).Font.Size = "11"
    '合并单元格
    .range("E1:G1").MergeCells = True
   
    '-------------------------------------------------标题区
    '项目信息
    .Cells(2, 2) = "墙厚"
    .Cells(2, 3) = "水平分布筋"
    .Cells(2, 6) = "配筋率"
    .Cells(2, 7) = "竖向分布筋"
    .Cells(2, 10) = "配筋率"
    
    .Cells(3, 3) = "直径"
    .Cells(3, 4) = "间隔"
    .Cells(3, 5) = "排数"

    .Cells(3, 7) = "直径"
    .Cells(3, 8) = "间隔"
    .Cells(3, 9) = "排数"
    
    '合并单元格
    .range("B2:B3").MergeCells = True
    .range("F2:F3").MergeCells = True
    .range("J2:J3").MergeCells = True

    .range("C2:E2").MergeCells = True
    .range("G2:I2").MergeCells = True
    
    
    '数据
    For i = 1 To 23
        .Cells(i + 3, 2) = 200 + (i - 1) * 50
    Next
    
End With

End Sub
