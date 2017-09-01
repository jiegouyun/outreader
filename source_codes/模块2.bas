Attribute VB_Name = "模块2"

Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/10/26
'1.新增对钢梁内力的提取


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%一层中所有柱或支撑柱标准内力提取
Sub BeamData_Y(Path, Beamsh, flonum, CorG As String)

Dim Beamsheet As String
Beamsheet = Beamsh
Dim flo As Integer
flo = flonum

Dim n As Integer
n = 0

Call Addsh(Beamsheet)

'清除工作表所有内容
Sheets(Beamsheet).Cells.Clear

With Sheets(Beamsheet)
   
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
    .Cells(1, 3) = "钢梁内力"
    .Cells(1, 13).Font.name = "黑体"
    .Cells(1, 3).Font.Size = "20"
    .Cells(1, 6) = flo & "F"
    '合并单元格
    .range("C1:E1").MergeCells = True
    '标题栏
    .Cells(2, 1) = "N-B"
    .Cells(2, 2) = "H"
    .Cells(2, 3) = "B1"
    .Cells(2, 4) = "B2"
    .Cells(2, 5) = "tw"
    .Cells(2, 6) = "tf1"
    .Cells(2, 7) = "tf2"
    .Cells(2, 8) = "(-M)"
    .Cells(2, 9) = "(+M)"
    .Cells(2, 10) = "Shear"
    
End With

'加背景色
Call AddShadow(Beamsheet, "A2:J2", 10092441)

'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim path1 As String, Filename As String, filepath1 As String, inputstring As String

'定义data为读入行的字符串
Dim data As String

'定义循环变量
Dim i, j As Integer

'定义构件编号变量
Dim mem As Integer

Dim FirstString_Beam As String


'==========================================================================================定义关键词变量

'柱或支撑编号行关键词
Dim Keyword_Beam As String
'赋值
Keyword_Beam = "N-B="



'==========================================================================================生成文件读取路径
'指定文件名为wpj_Num.out
Filename = "wpj" & flo & ".out"

'生成完整文件路径
filepath1 = Path & "\" & Filename


Sheets(Beamsheet).Select

Dim i_ As Integer: i = FreeFile

'打开结果文件
Open (filepath1) For Input Access Read As #i

'===========================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wpj.OUT"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #1, inputstring '读文本文件一行
  
    '将读取的一行字符串赋值与data变量
    data = inputstring

    '--------------------------------------------------------------------------定义各指标的判别字符
    '柱或支撑
     FirstString_Beam = Mid(data, 3, 4)
   
    '--------------------------------------------------------------------------读取柱或支撑的信息
    If FirstString_Beam = Keyword_Beam Then
       
        If CheckRegExpfromString(data, "B*H*U*T*D*F") = True Then
               
    '        Debug.Print FirstString_Beam, data
          
            '写入编号
            Sheets(Beamsheet).Cells(3 + n, 1) = StringfromStringforReg(data, "\d+\.?\d*", 1)
            '读取尺寸
            Sheets(Beamsheet).Cells(3 + n, 5) = StringfromStringforReg(data, "\d+\.?\d*", 5)
            Sheets(Beamsheet).Cells(3 + n, 2) = StringfromStringforReg(data, "\d+\.?\d*", 6)
            Sheets(Beamsheet).Cells(3 + n, 3) = StringfromStringforReg(data, "\d+\.?\d*", 7)
            Sheets(Beamsheet).Cells(3 + n, 6) = StringfromStringforReg(data, "\d+\.?\d*", 8)
            Sheets(Beamsheet).Cells(3 + n, 4) = StringfromStringforReg(data, "\d+\.?\d*", 9)
            Sheets(Beamsheet).Cells(3 + n, 7) = StringfromStringforReg(data, "\d+\.?\d*", 10)
   
            Do While Not EOF(1)
           
                Line Input #1, data
               
               '读取负弯矩
                If Mid(data, 3, 7) = "-M(kNm)" Then
                    For j = 1 To 9
                        Sheets(Beamsheet).Cells(3 + n, 10 + j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next
                    Sheets(Beamsheet).Cells(3 + n, 8) = "=MIN(RC[3]:RC[11])"
                End If
                
               '读取正弯矩
                If Mid(data, 3, 7) = "+M(kNm)" Then
                    For j = 1 To 9
                        Sheets(Beamsheet).Cells(3 + n, 19 + j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next
                    Sheets(Beamsheet).Cells(3 + n, 9) = "=MAX(RC[11]:RC[19])"
                End If
                
               '读取剪力
                If Mid(data, 3, 5) = "Shear" Then
                    For j = 1 To 9
                        Sheets(Beamsheet).Cells(3 + n, 28 + j) = StringfromStringforReg(data, "-?\d+\.?\d*", j)
                    Next
                    Sheets(Beamsheet).Cells(3 + n, 10) = "=MAX(RC[19]:RC[27])"
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
Call AddFormLine(Beamsheet, "A1:J" & Sheets(Beamsheet).range("A65535").End(xlUp).Row)

End Sub

Sub TEST()
Dim i
For i = 11 To 16
Sheets("Y_BEAM_F35").Columns(i).Delete
Next
End Sub
