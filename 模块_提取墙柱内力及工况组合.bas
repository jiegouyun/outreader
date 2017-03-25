Attribute VB_Name = "模块_提取墙柱内力及工况组合"
Option Explicit


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/20
'1.增加PKPM墙柱内力提取


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/23
'1.修正组合工况，增加竖向地震工况的读取和组合


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/14
'1.查看组合工况内力下，将工况最值高亮




'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
'                            YJK
'
'
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%一层中所有墙柱标准内力提取
Sub WallData_Y(Path, wallsh, flonum)

Dim wallsheet As String
wallsheet = wallsh


Dim n As Integer
n = 1

Dim flo As Integer
flo = flonum


Call Addsh(wallsheet)

'清除工作表所有内容
Sheets(wallsheet).Cells.Clear


''加表格线
'Call AddFormLine(wallsheet, "A2:H20000")

With Sheets(wallsheet)
   
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
    .Cells(1, 3) = "墙柱内力"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    
End With

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim path1 As String, Filename As String, filepath1 As String, inputstring As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

Dim FirstString_Wall As String


'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC ="



'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wwnl" & flo & ".out"

'生成完整文件路径
filepath1 = Path & "\" & Filename


Sheets(wallsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wwnl.OUT"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
  
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '墙
     FirstString_Wall = Mid(data, 2, 6)
   
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
       
        If StringfromStringforReg(data, "\d+\.?\d*", 1) = n Then
               
    '        Debug.Print FirstString_Wall, data
           
              '-------------------------------------------------标题区
            With Sheets(wallsheet)
            '项目信息
                .Cells(2 + (n - 1) * 17, 1) = "N-WC"
                .Cells(2 + (n - 1) * 17, 2) = "B"
                .Cells(2 + (n - 1) * 17, 3) = "H"
                .Cells(2 + (n - 1) * 17, 4) = "Lwc"
                .Cells(2 + (n - 1) * 17, 5) = "aa"
                .Cells(2 + (n - 1) * 17, 6) = "Angle"
                .Cells(2 + (n - 1) * 17, 7) = "Uc"
                
                .Cells(4 + (n - 1) * 17, 1) = "(iCase)"
                .Cells(4 + (n - 1) * 17, 2) = "Shear-X"
                .Cells(4 + (n - 1) * 17, 3) = "Shear-Y"
                .Cells(4 + (n - 1) * 17, 4) = "Axial"
                .Cells(4 + (n - 1) * 17, 5) = "Mx-Btm"
                .Cells(4 + (n - 1) * 17, 6) = "My-Btm"
                .Cells(4 + (n - 1) * 17, 7) = "Mx-Top"
                .Cells(4 + (n - 1) * 17, 8) = "My-Top"

            End With
            
            '加背景色
            Call AddShadow(wallsheet, "A" & 2 + (n - 1) * 17 & ":G" & 2 + (n - 1) * 17, 10092441)
            Call AddShadow(wallsheet, "A" & 4 + (n - 1) * 17 & ":H" & 4 + (n - 1) * 17, 10092441)
           
            '写入编号
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 1) = StringfromStringforReg(data, "\d+\.?\d*", 1)
            '读取角度
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 6) = StringfromStringforReg(data, "\d+\.?\d*", 5)
   
            Do While Not EOF(1)
           
                Line Input #1, data
               
                If Mid(data, 2, 9) = "*(    EX)" Then
                    Sheets(wallsheet).Cells(5 + (n - 1) * 17, 1) = "EX"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(5 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next
                
                End If
               
                If Mid(data, 2, 9) = "*(   EX+)" Then
                    Sheets(wallsheet).Cells(6 + (n - 1) * 17, 1) = "EX+"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(6 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
               
                If Mid(data, 2, 9) = "*(   EX-)" Then
                    Sheets(wallsheet).Cells(7 + (n - 1) * 17, 1) = "EX-"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(7 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(    EY)" Then
                    Sheets(wallsheet).Cells(8 + (n - 1) * 17, 1) = "EY"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(8 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   EY+)" Then
                    Sheets(wallsheet).Cells(9 + (n - 1) * 17, 1) = "EY+"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(9 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   EY-)" Then
                    Sheets(wallsheet).Cells(10 + (n - 1) * 17, 1) = "EY-"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(10 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   +WX)" Then
                    Sheets(wallsheet).Cells(11 + (n - 1) * 17, 1) = "WX+"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(11 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   -WX)" Then
                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, 1) = "WX-"
                    For j = 2 To 8
                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   +WY)" Then
                    Sheets(wallsheet).Cells(13 + (n - 1) * 17, 1) = "WY+"
                    For j = 2 To 8
                    Sheets(wallsheet).Cells(13 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   -WY)" Then
                    Sheets(wallsheet).Cells(14 + (n - 1) * 17, 1) = "WY-"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(14 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)

                    Next
                End If
                           
                If Mid(data, 2, 9) = "*(    DL)" Then
                    Sheets(wallsheet).Cells(15 + (n - 1) * 17, 1) = "DL"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(15 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(    LL)" Then
                    Sheets(wallsheet).Cells(16 + (n - 1) * 17, 1) = "LL"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(16 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                   
                If Mid(data, 2, 9) = "*(    EV)" Then
                    Sheets(wallsheet).Cells(17 + (n - 1) * 17, 1) = "EV"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(17 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                   
                If CheckRegExpfromString(data, "---") = True Then
                    Exit Do
                End If
                   
                Loop
               
                 n = n + 1
            End If
        End If

Loop

Close #1

'==========================================================================================读取构件尺寸等信息

'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String

'赋值
Keyword_Wall_UC = "Uc="

'柱、墙轴压比
Dim FirstString_Wall_UC As String

'序号归零
n = 1

'指定文件名为wpj_Num.out
Filename = "WPJ" & flo & ".OUT"

'生成完整文件路径
filepath1 = Path & "\" & Filename

i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历小震结果文件wpj" & flo; ".out; "
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
'        Debug.Print "读取" & flo & "层墙信息……"
        
        If StringfromStringforReg(data, "\d+", 1) = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 1) Then

            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 2) = StringfromStringforReg(data, "\d+\.?\d*", 4) * 1000
            B_w = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 2)
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 3) = StringfromStringforReg(data, "\d+\.?\d*", 5) * 1000
            H_w = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 3)
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 4) = StringfromStringforReg(data, "\d+\.?\d*", 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                FirstString_Wall_UC = Mid(data, 22, 3)
                If Mid(data, 3, 5) = "Cover" Then
                    Sheets(wallsheet).Cells(3 + (n - 1) * 17, 5) = StringfromStringforReg(data, "\d+", 2)
                End If
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
'                    Debug.Print "读取" & flo & "层墙轴压比……"
                    Sheets(wallsheet).Cells(3 + (n - 1) * 17, 7) = StringfromStringforReg(data, "0\.\d*", 1)
                End If
                If CheckRegExpfromString(data, "---") = True Then
                    Exit Do
                End If
            Loop
            
            n = n + 1
            
        End If
        
    End If

Loop

    
Close #1

Call AddFormLine(wallsheet, "A2:H" & Sheets(wallsheet).range("A65535").End(xlUp).Row)

End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%组合工况读取
Sub LOADCOMB_WC_Y(Path)

Dim lcombsh As String
lcombsh = "LCOMB_Y"

Call Addsh(lcombsh)

'清除工作表所有内容
Sheets(lcombsh).Cells.Clear


''加表格线
'Call AddFormLine(lcombsh, "A2:M20000")

'加背景色
Call AddShadow(lcombsh, "A2:M2", 10092441)

With Sheets(lcombsh)
   
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
    .Cells(1, 6) = "工况组合"
    .Cells(1, 6).Font.name = "黑体"
    .Cells(1, 6).Font.Size = "20"
    '合并单元格
    .range("F1:H1").MergeCells = True
    
End With

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer, j As Integer

Dim n As Integer
n = 1

'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_WC As String

'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wpj1.OUT"

'生成完整文件路径
filepath1 = Path & "\" & Filename

Sheets(lcombsh).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '--------------------------------------------------------------------------读取组合工况的信息
    If Mid(data, 6, 5) = "Ncm  " Then
    
        Debug.Print data
        
        For i = 1 To 13
            Sheets(lcombsh).Cells(2, i) = "(" & StringfromStringforReg(data, "\S+", i) & ")"
        Next
        
        Do While Not EOF(1)
            Line Input #1, inputstring '读文本文件一行
            data = inputstring
                
            If CheckRegExpfromString(data, "\*\*\*\*\*") = True Then
                Exit Do
            End If
            
'            If CheckRegExpfromString(data, "\d+") = True Then
            Debug.Print data
            If StringfromStringforReg(data, "\d+", 1) = n Then
                Debug.Print data
                '读取
                For i = 1 To 13
                    Sheets(lcombsh).Cells(n + 2, i) = StringfromStringforReg(data, "\S+", i)
                Next
                
                For i = 1 To 13
                    If Sheets(lcombsh).Cells(n + 2, i) = "--" Then Sheets(lcombsh).Cells(n + 2, i) = 0
                Next
                
            End If
               
            n = n + 1
'            End If
        Loop
    End If
Loop

Close #1

'加表格线
Call AddFormLine(lcombsh, "A2:M" & Sheets(lcombsh).range("A65535").End(xlUp).Row)

End Sub


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%单个构件内力、组合工况读取计算

Sub SingleWallData_Y(Path As String, sheetname As String, flonum As Integer, mem As Integer)


Dim wcsheet As String
wcsheet = sheetname


Dim n As Integer
n = 1

Dim i As Integer
Dim j As Integer

Dim flo As Integer
flo = flonum


Call Addsh(wcsheet)

'清除工作表所有内容
Sheets(wcsheet).Cells.Clear


'加表格线
'Call AddFormLine(wcsheet, "A2:H20000")
Call AddFormLine(wcsheet, "J4:P21")

'冻结首行首列
Sheets(wcsheet).Select
range("b5").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

With Sheets(wcsheet)
   
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
    .Columns("J:J").ColumnWidth = 18.13
   
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 3) = "墙柱校核"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    

    '项目信息
    .Cells(2, 1) = "N-WC"
    .Cells(2, 2) = "B"
    .Cells(2, 3) = "H"
    .Cells(2, 4) = "Lwc"
    .Cells(2, 5) = "aa"
    .Cells(2, 6) = "Angle"
    .Cells(2, 7) = "Uc"
    
    .Cells(4, 1) = "(iCase)"
    .Cells(4, 2) = "Shear-X"
    .Cells(4, 3) = "Shear-Y"
    .Cells(4, 4) = "Axial"
    .Cells(4, 5) = "Mx-Btm"
    .Cells(4, 6) = "My-Btm"
    .Cells(4, 7) = "Mx-Top"
    .Cells(4, 8) = "My-Top"
    
    .Cells(18, 1) = "(Ncm)"
    .Cells(18, 2) = "Shear-X"
    .Cells(18, 3) = "Shear-Y"
    .Cells(18, 4) = "Axial"
    .Cells(18, 5) = "Mx-Btm"
    .Cells(18, 6) = "My-Btm"
    .Cells(18, 7) = "Mx-Top"
    .Cells(18, 8) = "My-Top"
    
    .Cells(4, 10) = "LoadComb"
    .Cells(4, 11) = "Ncm"
    .Cells(4, 12) = "Shear-X"
    .Cells(4, 13) = "Shear-Y"
    .Cells(4, 14) = "Axial"
    .Cells(4, 15) = "Mx-Btm"
    .Cells(4, 16) = "My-Btm"
    
    '加背景色
    Call AddShadow(wcsheet, "A2:G2", 10092441)
    Call AddShadow(wcsheet, "A4:H4", 10092441)
    Call AddShadow(wcsheet, "A18:H18", 10092441)
    Call AddShadow(wcsheet, "J4:P4", 10092441)
    Call AddShadow(wcsheet, "J4:J21", 10092441)
    
End With

Dim m As Integer

'读取组合工况的个数
m = Sheets("LCOMB_Y").range("A65535").End(xlUp).Row - 2


Dim sh As Worksheet
'搜寻已有的工作表的名称
Dim condition

condition = "no"

For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = "Y_WCD_F" & flo Then
        condition = "yes"
    End If
Next

If condition = "no" Then
    MsgBox ("缺少表格Y_WCD_F" & flo)
    
    Exit Sub

End If


With Sheets(wcsheet)

    .Cells(3, 1) = mem
    
    '构件尺寸、轴压比等
    For i = 2 To 7
        .Cells(3, i).FormulaR1C1 = "=INDEX(Y_WCD_F" & flo & "!C,3+(R3C1-1)*17)"
    Next
    
    '构件标准内力
    For i = 4 To 17
        For j = 1 To 8
            .Cells(i, j).FormulaR1C1 = "=INDEX(Y_WCD_F" & flo & "!C," & i & "+(R3C1-1)*17)"
        Next
    Next
    
    '组合内力
    For i = 19 To m + 18
        For j = 2 To 8
            .Cells(i, j).FormulaR1C1 = "=R15C*LCOMB_Y!R[-16]C2+R16C*LCOMB_Y!R[-16]C3+R11C*LCOMB_Y!R[-16]C4+R12C*LCOMB_Y!R[-16]C5+R13C*LCOMB_Y!R[-16]C6+R14C*LCOMB_Y!R[-16]C7+R5C*LCOMB_Y!R[-16]C8+R8C*LCOMB_Y!R[-16]C9+R17C*LCOMB_Y!R[-16]C10"
        Next
        .Cells(i, 1) = i - 18
    Next
    
    '最不利工况
    For i = 5 To 18
        For j = 12 To 16
            .Cells(i, j).FormulaR1C1 = "=VLOOKUP(RC11,R18C1:R" & m + 17 & "C8," & j - 10 & ")"
        Next
    Next
    
    For i = 12 To 16
        .Cells(19, i).FormulaR1C1 = "=1.2*R15C[-10]+1.4*R16C[-10]"
    Next
    
    For i = 12 To 16
        .Cells(20, i).FormulaR1C1 = "=R15C[-10]+R16C[-10]"
    Next
    
    .Cells(21, 10) = "Nmin"
    .Cells(21, 11) = "=MAX(R[-2]C[-7]:R[100]C[-7])"
    
    
    Dim ii As Integer
    Dim i_RowID As Integer
    Dim i_Rng As range
    
    '---------------------------------------------------------工况组合
    For ii = 2 To 8
    Dim R As range
    Set R = Worksheets(sheetname).range(Cells(18, ii), Cells(m + 17, ii))
    Call maxormin(R, "max", sheetname & "!R18C" & CStr(ii) & ":R" & CStr(m + 17) & "C" & CStr(ii))
    Call maxormin(R, "min", sheetname & "!R18C" & CStr(ii) & ":R" & CStr(m + 17) & "C" & CStr(ii))
    Next

    
    
End With



'==========================================================================================生成文件读取路径
'指定文件名为wdcnl.OUT
Dim Filename As String, filepath1 As String

Filename = "wdcnl.OUT"

'生成完整文件路径
filepath1 = Path & "\" & Filename

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "读取相关指标"
Debug.Print "……"

Dim inputstring As String, data As String

n = 5

Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '--------------------------------------------------------------------------读取组合工况的信息
    If Mid(data, 2, 8) = "N-WC =1 " Then
    
        Debug.Print data
        
        
        Do While Not EOF(1)
            Line Input #1, inputstring '读文本文件一行
            data = inputstring
                
            If CheckRegExpfromString(data, "----") = True Then
                Exit Do
            End If
            
'            If CheckRegExpfromString(data, "\d+") = True Then
'            Debug.Print data
            
            Sheets(wcsheet).Cells(n, 11) = StringfromStringforReg(data, "\d+", 1)
            Sheets(wcsheet).Cells(n, 10) = "(" & StringfromStringforReg(data, "\S+", 8) & ")"
            n = n + 1
        
        Loop
    End If
Loop

Close #1

Call AddFormLine(wcsheet, "A2:H" & Sheets(wcsheet).range("A65535").End(xlUp).Row)

End Sub






'============================================================================================================================================================================
'============================================================================================================================================================================
'============================================================================================================================================================================






'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%                                                                        %%
'                            PKPM
'%%                                                                        %%
'%%                                                                        %%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%一层中所有墙柱标准内力提取
Sub WallData_P(Path, wallsh, flonum)

Dim wallsheet As String
wallsheet = wallsh


Dim n As Integer
n = 1

Dim flo As Integer
flo = flonum


Call Addsh(wallsheet)

'清除工作表所有内容
Sheets(wallsheet).Cells.Clear


''加表格线
'Call AddFormLine(wallsheet, "A2:H20000")

With Sheets(wallsheet)
   
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
    .Cells(1, 3) = "墙柱内力"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    
End With

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim path1 As String, Filename As String, filepath1 As String, inputstring As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

Dim FirstString_Wall As String


'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-Wc ="


'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wwnl" & flo & ".out"

'生成完整文件路径
filepath1 = Path & "\" & Filename


Sheets(wallsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wwnl.OUT"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
  
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '墙
     FirstString_Wall = Mid(data, 2, 6)
   
    '--------------------------------------------------------------------------读取墙的信息
    If FirstString_Wall = Keyword_Wall Then
       
        If StringfromStringforReg(data, "\d+\.?\d*", 1) = n Then
               
    '        Debug.Print FirstString_Wall, data
           
              '-------------------------------------------------标题区
            With Sheets(wallsheet)
            '项目信息
                .Cells(2 + (n - 1) * 17, 1) = "N-WC"
                .Cells(2 + (n - 1) * 17, 2) = "B"
                .Cells(2 + (n - 1) * 17, 3) = "H"
                .Cells(2 + (n - 1) * 17, 4) = "Lwc"
                .Cells(2 + (n - 1) * 17, 5) = "aa"
                .Cells(2 + (n - 1) * 17, 6) = "Angle"
                .Cells(2 + (n - 1) * 17, 7) = "Uc"
                
                .Cells(4 + (n - 1) * 17, 1) = "(iCase)"
                .Cells(4 + (n - 1) * 17, 2) = "Shear-X"
                .Cells(4 + (n - 1) * 17, 3) = "Shear-Y"
                .Cells(4 + (n - 1) * 17, 4) = "Axial"
                .Cells(4 + (n - 1) * 17, 5) = "Mx-Btm"
                .Cells(4 + (n - 1) * 17, 6) = "My-Btm"
                .Cells(4 + (n - 1) * 17, 7) = "Mx-Top"
                .Cells(4 + (n - 1) * 17, 8) = "My-Top"

            End With
            
            '加背景色
            Call AddShadow(wallsheet, "A" & 2 + (n - 1) * 17 & ":G" & 2 + (n - 1) * 17, 10092441)
            Call AddShadow(wallsheet, "A" & 4 + (n - 1) * 17 & ":H" & 4 + (n - 1) * 17, 10092441)
           
            '写入编号
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 1) = StringfromStringforReg(data, "\d+\.?\d*", 1)
            '读取角度
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 6) = StringfromStringforReg(data, "\d+\.?\d*", 5)
   
            Do While Not EOF(1)
           
                Line Input #1, data
               
                If Mid(data, 2, 5) = "( 1*)" Then
                    Sheets(wallsheet).Cells(5 + (n - 1) * 17, 1) = "EX"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(5 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next
                
                End If
               
                If Mid(data, 2, 5) = "( 2*)" Then
                    Sheets(wallsheet).Cells(6 + (n - 1) * 17, 1) = "EX+"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(6 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
               
                If Mid(data, 2, 5) = "( 3*)" Then
                    Sheets(wallsheet).Cells(7 + (n - 1) * 17, 1) = "EX-"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(7 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
                If Mid(data, 2, 5) = "( 4*)" Then
                    Sheets(wallsheet).Cells(8 + (n - 1) * 17, 1) = "EY"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(8 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
                If Mid(data, 2, 5) = "( 5*)" Then
                    Sheets(wallsheet).Cells(9 + (n - 1) * 17, 1) = "EY+"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(9 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
                If Mid(data, 2, 5) = "( 6*)" Then
                    Sheets(wallsheet).Cells(10 + (n - 1) * 17, 1) = "EY-"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(10 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
                If Mid(data, 2, 5) = "( 7 )" Then
                    Sheets(wallsheet).Cells(11 + (n - 1) * 17, 1) = "WX"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(11 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
'                If Mid(data, 2, 5) = "( 7 )" Then
'                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, 1) = "WX-"
'                    For j = 2 To 8
'                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
'                    Next
'
'                End If
                           
                If Mid(data, 2, 5) = "( 8 )" Then
                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, 1) = "WY"
                    For j = 2 To 8
                    Sheets(wallsheet).Cells(12 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
'                If Mid(data, 2, 5) = "( 8 )" Then
'                    Sheets(wallsheet).Cells(14 + (n - 1) * 17, 1) = "WY-"
'                    For j = 2 To 8
'                        Sheets(wallsheet).Cells(14 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
'
'                    Next
'                End If
                           
                If Mid(data, 2, 5) = "( 9 )" Then
                    Sheets(wallsheet).Cells(13 + (n - 1) * 17, 1) = "DL"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(13 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                           
                If Mid(data, 2, 5) = "(10 )" Then
                    Sheets(wallsheet).Cells(14 + (n - 1) * 17, 1) = "LL"
                    For j = 2 To 8
                        Sheets(wallsheet).Cells(14 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next

                End If
                   
'                If Mid(data, 2, 9) = "*(    EV)" Then
'                    Sheets(wallsheet).Cells(17 + (n - 1) * 17, 1) = "EV"
'                    For j = 2 To 8
'                        Sheets(wallsheet).Cells(17 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
'                    Next
'
'                End If
                   
                If CheckRegExpfromString(data, "---") = True Then
                    Exit Do
                End If
                   
                Loop
               
                 n = n + 1
            End If
        End If

Loop

Close #1

'==========================================================================================读取构件尺寸等信息

'赋值
Keyword_Wall = "N-WC="

'柱、墙轴压比行关键词
Dim Keyword_Wall_UC As String

'赋值
Keyword_Wall_UC = "Uc="

'柱、墙轴压比
Dim FirstString_Wall_UC As String

'序号归零
n = 1

'指定文件名为wpj_Num.out
Filename = "WPJ" & flo & ".OUT"

'生成完整文件路径
filepath1 = Path & "\" & Filename

i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历小震结果文件wpj" & flo; ".out; "
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
'        Debug.Print "读取" & flo & "层墙信息……"
        
        If StringfromStringforReg(data, "\d+", 1) = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 1) Then

            '读取墙截面
            Dim B_w As Long, H_w As Long
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 2) = StringfromStringforReg(data, "\d+\.?\d*", 4) * 1000
            B_w = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 2)
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 3) = StringfromStringforReg(data, "\d+\.?\d*", 5) * 1000
            H_w = Sheets(wallsheet).Cells(3 + (n - 1) * 17, 3)
            Sheets(wallsheet).Cells(3 + (n - 1) * 17, 4) = StringfromStringforReg(data, "\d+\.?\d*", 6) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                FirstString_Wall_UC = Mid(data, 22, 3)
                If Mid(data, 3, 5) = "Cover" Then
                    Sheets(wallsheet).Cells(3 + (n - 1) * 17, 5) = StringfromStringforReg(data, "\d+", 2)
                End If
                If FirstString_Wall_UC = Keyword_Wall_UC Then
                    '读取墙轴压比
'                    Debug.Print "读取" & flo & "层墙轴压比……"
                    Sheets(wallsheet).Cells(3 + (n - 1) * 17, 7) = StringfromStringforReg(data, "0\.\d*", 1)
                End If
                If CheckRegExpfromString(data, "---") = True Then
                    Exit Do
                End If
            Loop
            
            n = n + 1
            
        End If
        
    End If

Loop

    
Close #1

Call AddFormLine(wallsheet, "A2:H" & Sheets(wallsheet).range("A65535").End(xlUp).Row)

End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%组合工况读取
Sub LOADCOMB_WC_P(Path)

On Error Resume Next

Dim lcombsh As String
lcombsh = "LCOMB_P"

Call Addsh(lcombsh)

'清除工作表所有内容
Sheets(lcombsh).Cells.Clear


''加表格线
'Call AddFormLine(lcombsh, "A2:M20000")

'加背景色
Call AddShadow(lcombsh, "A2:M2", 10092441)

With Sheets(lcombsh)
   
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
    .Cells(1, 6) = "工况组合"
    .Cells(1, 6).Font.name = "黑体"
    .Cells(1, 6).Font.Size = "20"
    '合并单元格
    .range("F1:H1").MergeCells = True
    
End With

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath1, inputstring   As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i As Integer, j As Integer

Dim n As Integer
n = 1

'==========================================================================================定义关键词变量

'墙编号行关键词
Dim Keyword_Wall As String
'赋值
Keyword_Wall = "N-WC="

'==========================================================================================定义首字符变量

'柱、墙、梁
Dim FirstString_WC As String

'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "WPJ1.OUT"

'生成完整文件路径
filepath1 = Path & "\" & Filename

Sheets(lcombsh).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
   
    '将读取的一行字符串赋值与data变量
    data = inputstring
    Debug.Print data
    
    '--------------------------------------------------------------------------读取组合工况的信息
    If Mid(data, 2, 3) = "Ncm" Then
    
        Debug.Print data
        
        For i = 1 To 13
            Sheets(lcombsh).Cells(2, i) = "(" & StringfromStringforReg(data, "\S+", i) & ")"
        Next
        
        Do While Not EOF(1)
            Line Input #1, inputstring '读文本文件一行
            data = inputstring
                
            If CheckRegExpfromString(data, "-------") = True Then
                Exit Do
            End If
            
'            If CheckRegExpfromString(data, "\d+") = True Then
            Debug.Print data
            If StringfromStringforReg(data, "\d+", 1) = n Then
                Debug.Print data
                '读取
                For i = 1 To 13
                    Sheets(lcombsh).Cells(n + 2, i) = StringfromStringforReg(data, "\S+", i)
                Next
                
                For i = 1 To 13
                    If Sheets(lcombsh).Cells(n + 2, i) = "--" Then Sheets(lcombsh).Cells(n + 2, i) = 0
                Next
                
            End If
               
            n = n + 1
'            End If
        Loop
    End If
Loop

Close #1

'加表格线
Call AddFormLine(lcombsh, "A2:M" & Sheets(lcombsh).range("A65535").End(xlUp).Row)

End Sub


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%单个构件内力、组合工况读取计算

Sub SingleWallData_P(Path As String, sheetname As String, flonum As Integer, mem As Integer)


Dim wcsheet As String
wcsheet = sheetname


Dim n As Integer
n = 1

Dim i As Integer
Dim j As Integer

Dim flo As Integer
flo = flonum


Call Addsh(wcsheet)

'清除工作表所有内容
Sheets(wcsheet).Cells.Clear


'加表格线
'Call AddFormLine(wcsheet, "A2:H20000")
Call AddFormLine(wcsheet, "J4:P21")

'冻结首行首列
Sheets(wcsheet).Select
range("b5").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

With Sheets(wcsheet)
   
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
    .Columns("J:J").ColumnWidth = 18.13
   
    '-------------------------------------------------表头区
    '表头
    .Cells(1, 3) = "墙柱校核"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    

    '项目信息
    .Cells(2, 1) = "N-WC"
    .Cells(2, 2) = "B"
    .Cells(2, 3) = "H"
    .Cells(2, 4) = "Lwc"
    .Cells(2, 5) = "aa"
    .Cells(2, 6) = "Angle"
    .Cells(2, 7) = "Uc"
    
    .Cells(4, 1) = "(iCase)"
    .Cells(4, 2) = "Shear-X"
    .Cells(4, 3) = "Shear-Y"
    .Cells(4, 4) = "Axial"
    .Cells(4, 5) = "Mx-Btm"
    .Cells(4, 6) = "My-Btm"
    .Cells(4, 7) = "Mx-Top"
    .Cells(4, 8) = "My-Top"
    
    .Cells(18, 1) = "(Ncm)"
    .Cells(18, 2) = "Shear-X"
    .Cells(18, 3) = "Shear-Y"
    .Cells(18, 4) = "Axial"
    .Cells(18, 5) = "Mx-Btm"
    .Cells(18, 6) = "My-Btm"
    .Cells(18, 7) = "Mx-Top"
    .Cells(18, 8) = "My-Top"
    
    .Cells(4, 10) = "LoadComb"
    .Cells(4, 11) = "Ncm"
    .Cells(4, 12) = "Shear-X"
    .Cells(4, 13) = "Shear-Y"
    .Cells(4, 14) = "Axial"
    .Cells(4, 15) = "Mx-Btm"
    .Cells(4, 16) = "My-Btm"
    
    '加背景色
    Call AddShadow(wcsheet, "A2:G2", 10092441)
    Call AddShadow(wcsheet, "A4:H4", 10092441)
    Call AddShadow(wcsheet, "A18:H18", 10092441)
    Call AddShadow(wcsheet, "J4:P4", 10092441)
    Call AddShadow(wcsheet, "J4:J21", 10092441)
    
End With

Dim m As Integer

'读取组合工况的个数
m = Sheets("LCOMB_P").range("A65535").End(xlUp).Row - 2


Dim sh As Worksheet
'搜寻已有的工作表的名称
Dim condition

condition = "no"

For Each sh In Worksheets
    '如果与新定义的工作表名相同，则退出程序
    If sh.name = "P_WCD_F" & flo Then
        condition = "yes"
    End If
Next

If condition = "no" Then
    MsgBox ("缺少表格P_WCD_F" & flo)
    
    Exit Sub

End If


With Sheets(wcsheet)

    .Cells(3, 1) = mem
    
    '构件尺寸、轴压比等
    For i = 2 To 7
        .Cells(3, i).FormulaR1C1 = "=INDEX(P_WCD_F" & flo & "!C,3+(R3C1-1)*17)"
    Next
    
    '构件标准内力
    For i = 4 To 17
        For j = 1 To 8
            .Cells(i, j).FormulaR1C1 = "=INDEX(P_WCD_F" & flo & "!C," & i & "+(R3C1-1)*17)"
        Next
    Next
    
    '组合内力
    For i = 19 To m + 18
        For j = 2 To 8
            .Cells(i, j).FormulaR1C1 = "=R13C*LCOMB_P!R[-16]C2+R14C*LCOMB_P!R[-16]C3+R11C*LCOMB_P!R[-16]C4+R12C*LCOMB_P!R[-16]C5+R5C*LCOMB_P!R[-16]C6+R8C*LCOMB_P!R[-16]C7"
        Next
        .Cells(i, 1) = i - 18
    Next
    
    '最不利工况
    For i = 5 To 18
        For j = 12 To 16
            .Cells(i, j).FormulaR1C1 = "=VLOOKUP(RC11,R18C1:R" & m + 17 & "C8," & j - 10 & ")"
        Next
    Next
    

Sheets(wcsheet).Cells(19, 10) = "1.2*DL+1.4*LL"
Sheets(wcsheet).Cells(20, 10) = "DL+LL"

    For i = 12 To 16
        .Cells(19, i).FormulaR1C1 = "=1.2*R13C[-10]+1.4*R14C[-10]"
    Next
    
    For i = 12 To 16
        .Cells(20, i).FormulaR1C1 = "=R13C[-10]+R14C[-10]"
    Next
    
    .Cells(21, 10) = "Nmin"
    .Cells(21, 11) = "=MAX(R[-2]C[-7]:R[100]C[-7])"
    
    
    Dim ii As Integer
    Dim i_RowID As Integer
    Dim i_Rng As range
    
    '---------------------------------------------------------工况组合
    For ii = 2 To 8
    Dim R As range
    Set R = Worksheets(sheetname).range(Cells(18, ii), Cells(m + 17, ii))
    Call maxormin(R, "max", sheetname & "!R18C" & CStr(ii) & ":R" & CStr(m + 17) & "C" & CStr(ii))
    Call maxormin(R, "min", sheetname & "!R18C" & CStr(ii) & ":R" & CStr(m + 17) & "C" & CStr(ii))
    Next

    
    
End With


'
''==========================================================================================生成文件读取路径
''指定文件名为wdcnl.OUT
'Dim Filename As String, filepath1 As String
'
'Filename = "wdcnl.OUT"
'
''生成完整文件路径
'filepath1 = Path & "\" & Filename
'
'Dim i_ As Integer: i = FreeFile
'
''打开结果文件
'Open (filepath1) For Input Access Read As #i
'
''===========================================================================================逐行读取文本
'
'Debug.Print "读取相关指标"
'Debug.Print "……"
'
'Dim inputstring As String, data As String
'
'n = 5
'
'Do While Not EOF(1)
'
'    Line Input #1, inputstring '读文本文件一行
'
'    '将读取的一行字符串赋值与data变量
'    data = inputstring
'
'    '--------------------------------------------------------------------------读取组合工况的信息
'    If Mid(data, 2, 8) = "N-WC =1 " Then
'
'        Debug.Print data
'
'
'        Do While Not EOF(1)
'            Line Input #1, inputstring '读文本文件一行
'            data = inputstring
'
'            If CheckRegExpfromString(data, "----") = True Then
'                Exit Do
'            End If
'
''            If CheckRegExpfromString(data, "\d+") = True Then
''            Debug.Print data
'
'            Sheets(wcsheet).Cells(n, 11) = StringfromStringforReg(data, "\d+", 1)
'            Sheets(wcsheet).Cells(n, 10) = "(" & StringfromStringforReg(data, "\S+", 8) & ")"
'            n = n + 1
'
'        Loop
'    End If
'Loop
'
'Close #1


For n = 5 To 18
  Sheets(wcsheet).Cells(n, 11) = 1
Next


Call AddFormLine(wcsheet, "A2:H" & Sheets(wcsheet).range("A65535").End(xlUp).Row)

End Sub






