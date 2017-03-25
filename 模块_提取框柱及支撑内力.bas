Attribute VB_Name = "模块_提取框柱及支撑内力"

Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/7/05
'1.修改提取支撑内力时的错误


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/23
'1.添加框柱、支撑内力提取


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%一层中所有柱或支撑柱标准内力提取
Sub ColData_Y(Path, Colsh, flonum, CorG As String)

Dim Colsheet As String
Colsheet = Colsh
Dim flo As Integer
flo = flonum

Dim n As Integer
n = 1

Call Addsh(Colsheet)

'清除工作表所有内容
Sheets(Colsheet).Cells.Clear

With Sheets(Colsheet)
   
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
    .Cells(1, 3) = "柱或支撑内力"
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

Dim FirstString_Col As String


'==========================================================================================定义关键词变量

'柱或支撑编号行关键词
Dim Keyword_Col As String
'赋值
Keyword_Col = "N-" & CorG & " ="



'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wwnl" & flo & ".out"

'生成完整文件路径
filepath1 = Path & "\" & Filename


Sheets(Colsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wdcnl.OUT"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
  
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱或支撑
     FirstString_Col = Mid(data, 2, 5)
   
    '--------------------------------------------------------------------------读取柱或支撑的信息
    If FirstString_Col = Keyword_Col Then
       
        If StringfromStringforReg(data, "\d+\.?\d*", 1) = n Then
               
    '        Debug.Print FirstString_Col, data
           
              '-------------------------------------------------标题区
            With Sheets(Colsheet)
            '项目信息
                .Cells(2 + (n - 1) * 17, 1) = "N-" & CorG
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
            Call AddShadow(Colsheet, "A" & 2 + (n - 1) * 17 & ":G" & 2 + (n - 1) * 17, 10092441)
            Call AddShadow(Colsheet, "A" & 4 + (n - 1) * 17 & ":H" & 4 + (n - 1) * 17, 10092441)
           
            '写入编号
            Sheets(Colsheet).Cells(3 + (n - 1) * 17, 1) = StringfromStringforReg(data, "\d+\.?\d*", 1)
            '读取角度
            Sheets(Colsheet).Cells(3 + (n - 1) * 17, 6) = StringfromStringforReg(data, "\d+\.?\d*", 5)
   
            Do While Not EOF(1)
           
                Line Input #1, data
               
                If Mid(data, 2, 9) = "*(    EX)" Then
                    Sheets(Colsheet).Cells(5 + (n - 1) * 17, 1) = "EX"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(5 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next
                
                End If
               
                If Mid(data, 2, 9) = "*(   EX+)" Then
                    Sheets(Colsheet).Cells(6 + (n - 1) * 17, 1) = "EX+"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(6 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
               
                If Mid(data, 2, 9) = "*(   EX-)" Then
                    Sheets(Colsheet).Cells(7 + (n - 1) * 17, 1) = "EX-"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(7 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(    EY)" Then
                    Sheets(Colsheet).Cells(8 + (n - 1) * 17, 1) = "EY"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(8 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   EY+)" Then
                    Sheets(Colsheet).Cells(9 + (n - 1) * 17, 1) = "EY+"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(9 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   EY-)" Then
                    Sheets(Colsheet).Cells(10 + (n - 1) * 17, 1) = "EY-"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(10 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   +WX)" Then
                    Sheets(Colsheet).Cells(11 + (n - 1) * 17, 1) = "WX+"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(11 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   -WX)" Then
                    Sheets(Colsheet).Cells(12 + (n - 1) * 17, 1) = "WX-"
                    For j = 2 To 8
                    Sheets(Colsheet).Cells(12 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   +WY)" Then
                    Sheets(Colsheet).Cells(13 + (n - 1) * 17, 1) = "WY+"
                    For j = 2 To 8
                    Sheets(Colsheet).Cells(13 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(   -WY)" Then
                    Sheets(Colsheet).Cells(14 + (n - 1) * 17, 1) = "WY-"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(14 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)

                    Next
                End If
                           
                If Mid(data, 2, 9) = "*(    DL)" Then
                    Sheets(Colsheet).Cells(15 + (n - 1) * 17, 1) = "DL"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(15 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                           
                If Mid(data, 2, 9) = "*(    LL)" Then
                    Sheets(Colsheet).Cells(16 + (n - 1) * 17, 1) = "LL"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(16 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
                    Next

                End If
                   
                If Mid(data, 2, 9) = "*(    EV)" Then
                    Sheets(Colsheet).Cells(17 + (n - 1) * 17, 1) = "EV"
                    For j = 2 To 8
                        Sheets(Colsheet).Cells(17 + (n - 1) * 17, j) = StringfromStringforReg(data, "-?\d+\.?\d*", j - 1)
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
Keyword_Col = "N-C="

'柱、柱或支撑轴压比行关键词
Dim Keyword_Col_UC As String

'赋值
Keyword_Col_UC = "Uc="

'柱、柱或支撑轴压比
Dim FirstString_Col_UC As String

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
    '柱、柱或支撑
    FirstString_Col = Mid(data, 3, 5)
    
    '--------------------------------------------------------------------------读取柱或支撑的信息
    If FirstString_Col = Keyword_Col Then
'        Debug.Print "读取" & flo & "层柱或支撑信息……"
        
        If StringfromStringforReg(data, "\d+", 1) = Sheets(Colsheet).Cells(3 + (n - 1) * 17, 1) Then

            '读取柱或支撑截面
            Dim B_w As Long, H_w As Long
            Sheets(Colsheet).Cells(3 + (n - 1) * 17, 2) = StringfromStringforReg(data, "\d+\.?\d*", 3) * 1000
            B_w = Sheets(Colsheet).Cells(3 + (n - 1) * 17, 2)
            Sheets(Colsheet).Cells(3 + (n - 1) * 17, 3) = StringfromStringforReg(data, "\d+\.?\d*", 4) * 1000
            H_w = Sheets(Colsheet).Cells(3 + (n - 1) * 17, 3)
            Sheets(Colsheet).Cells(3 + (n - 1) * 17, 4) = StringfromStringforReg(data, "\d+\.?\d*", 5) * 1000

            Do While Not EOF(1)
                Line Input #1, data
                FirstString_Col_UC = Mid(data, 22, 3)
                If Mid(data, 3, 5) = "Cover" Then
                    Sheets(Colsheet).Cells(3 + (n - 1) * 17, 5) = StringfromStringforReg(data, "\d+", 2)
                End If
                If FirstString_Col_UC = Keyword_Col_UC Then
                    '读取柱或支撑轴压比
'                    Debug.Print "读取" & flo & "层柱或支撑轴压比……"
                    Sheets(Colsheet).Cells(3 + (n - 1) * 17, 7) = StringfromStringforReg(data, "0\.\d*", 1)
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

'加表格线
Call AddFormLine(Colsheet, "A2:H" & Sheets(Colsheet).range("A65535").End(xlUp).Row)

End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%组合工况读取
Sub LOADCOMB_C_Y(Path)

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

'柱或支撑编号行关键词
Dim Keyword_Col As String
'赋值
Keyword_Col = "N-" & NorG & "="

'==========================================================================================定义首字符变量

'柱、柱或支撑、梁
Dim FirstString_C As String

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

Sub SingleColData_Y(Path As String, sheetname As String, flonum As Integer, mem As Integer, CorG As String)


Dim Ccsheet As String
Ccsheet = sheetname


Dim n As Integer
n = 1

Dim i As Integer
Dim j As Integer

Dim flo As Integer
flo = flonum


Call Addsh(Ccsheet)

'清除工作表所有内容
Sheets(Ccsheet).Cells.Clear


'加表格线
'Call AddFormLine(Ccsheet, "A2:H20000")
Call AddFormLine(Ccsheet, "J4:P21")

'冻结首行首列
Sheets(Ccsheet).Select
range("b5").Select
With ActiveWindow
    .SplitColumn = 1
    .SplitRow = 3
End With
ActiveWindow.FreezePanes = True

With Sheets(Ccsheet)
   
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
    .Cells(1, 3) = "柱或支撑校核"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    

    '项目信息
    .Cells(2, 1) = "N-" & CorG
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
    Call AddShadow(Ccsheet, "A2:G2", 10092441)
    Call AddShadow(Ccsheet, "A4:H4", 10092441)
    Call AddShadow(Ccsheet, "A18:H18", 10092441)
    Call AddShadow(Ccsheet, "J4:P4", 10092441)
    Call AddShadow(Ccsheet, "J4:J21", 10092441)
    
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
    If sh.name = "Y_" & CorG & "D_F" & flo Then
        condition = "yes"
    End If
Next

If condition = "no" Then
    MsgBox ("缺少表格Y_" & CorG & "D_F" & flo)
    
    Exit Sub

End If

With Sheets(Ccsheet)

    .Cells(3, 1) = mem
    
    '构件尺寸、轴压比等
    For i = 2 To 7
        .Cells(3, i).FormulaR1C1 = "=INDEX(Y_" & CorG & "D_F" & flo & "!C,3+(R3C1-1)*17)"
    Next
    
    '构件标准内力
    For i = 4 To 17
        For j = 1 To 8
            .Cells(i, j).FormulaR1C1 = "=INDEX(Y_" & CorG & "D_F" & flo & "!C," & i & "+(R3C1-1)*17)"
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
            .Cells(i, j).FormulaR1C1 = "=VLOOKUP(RC11,R19C1:R" & m + 18 & "C8," & j - 10 & ")"
        Next
    Next
    
    For i = 12 To 16
        .Cells(19, i).FormulaR1C1 = "=1.2*R15C[-10]+1.4*R16C[-10]"
    Next
    
    For i = 12 To 16
        .Cells(20, i).FormulaR1C1 = "=R15C[-10]+R16C[-10]"
    Next
    
    .Cells(21, 10) = "Nmin"
    .Cells(21, 11) = "=MAX(R[-4]C[-7]:R[30]C[-7])"
    
    
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
    If Mid(data, 2, 7) = "N-C =1 " Then
    
        Debug.Print data
        
        
        Do While Not EOF(1)
            Line Input #1, inputstring '读文本文件一行
            data = inputstring
                
            If CheckRegExpfromString(data, "----") = True Then
                Exit Do
            End If
            
'            If CheckRegExpfromString(data, "\d+") = True Then
'            Debug.Print data
            
            Sheets(Ccsheet).Cells(n, 11) = StringfromStringforReg(data, "\d+", 1)
            Sheets(Ccsheet).Cells(n, 10) = "(" & StringfromStringforReg(data, "\S+", 8) & ")"
            n = n + 1
        
        Loop
    End If
Loop

For n = 5 To 18
    Sheets(Ccsheet).Cells(n, 11) = 1
Next

Close #1

'加表格线
Call AddFormLine(Ccsheet, "A2:H" & Sheets(Ccsheet).range("A65535").End(xlUp).Row)

End Sub

