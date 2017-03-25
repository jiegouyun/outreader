Attribute VB_Name = "MBuilding_侧向刚度"

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/05/20
'更新内容:
'1.添加判断，修正刚度比提取
'2.Midas输出的刚度比版本不同有不一致的地方，旧版为说明里的剪力/位移角，新版感觉又成了剪力/位移，对比时当极其注意。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15
'1.修改高亮代码，换为条件格式
'2.对刚度特殊位置进行高亮


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29
'1.增加高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/18

'更新内容：
'1.


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****            MBuilding_楼层侧向刚度验算.TXT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************



Sub OUTReader_MBuilding_侧向刚度(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename, filepath, inputstring  As String

'定义data为读入行的字符串
Dim data As String


'定义循环变量
Dim i, j As Integer

Dim ii As Integer




'==========================================================================================定义关键词变量


'刚度比
Dim Keyword_Rate1, Keyword_Rate2   As String
'赋值
Keyword_Rate1 = "RS_0"
Keyword_Rate2 = "RS_90"



'==========================================================================================定义首字符变量


'刚度比
Dim FirstString_kRate1 As String
Dim FirstString_kRate2 As String

'=============================================================================================================================生成文件读取路径

'指定文件名为wmass.out
Filename = Dir(Path & "\*_楼层侧向刚度验算.txt")

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件：楼层侧向刚度验算.TXT"
Debug.Print "读取相关指标"
Debug.Print "……"

ii = 0
Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行

    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符

    '刚度比
    FirstString_kRate1 = Mid(data, 10, 4)
    FirstString_kRate2 = Mid(data, 10, 5)
       

    '-------------------------------------------------------------------------------------------读取刚度比
    
        
    If FirstString_kRate1 = Keyword_Rate1 Then
        'Debug.Print "aaaaaaaaaaaaaa"
        Line Input #i, data
        Line Input #i, data
        
        Dim str As String
        Dim m, i2 As Integer
        For m = 1 To 6
        
        With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\S+"
        Dim mc
        Set mc = .Execute(data)  '执行匹配项查找
        If mc.Count >= m Then
            str = mc(m - 1).Value
        End If
        End With
    
'        str = StringfromStringforReg(data, "\S+", m)
        If str = "Rat1" Then
        i2 = m
        End If
        Next
        
        
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "Base") = True Then
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                Sheets("d_M").Cells(j, 2) = StringfromStringforReg(data, "\S+", i2)
                Sheets("d_M").Cells(j, 4) = StringfromStringforReg(data, "\S+", 3)
            End If
            If CheckRegExpfromString(data, "--") = True Then
                Exit Do
            End If
        Loop

    End If
    If FirstString_kRate2 = Keyword_Rate2 Then
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "Base") = True Then
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If
                Sheets("d_M").Cells(j, 3) = StringfromStringforReg(data, "\S+", i2)
                Sheets("d_M").Cells(j, 5) = StringfromStringforReg(data, "\S+", 3)
            End If
            If CheckRegExpfromString(data, "--") = True Then
                Exit Do
            End If
            
            ii = ii + 1
                       
        Loop
        'Num_all = ii
        'Debug.Print Num_All

    End If
    
          
Loop

Sheets("d_M").Cells(Num_all + 2, 2) = 1
Sheets("d_M").Cells(Num_all + 2, 3) = 1

'Sheets("d_M").Cells(Num_all + 2, 4) = Sheets("d_M").Cells(Num_all + 1, 4)
'Sheets("d_M").Cells(Num_all + 2, 5) = Sheets("d_M").Cells(Num_all + 1, 5)

'关闭结果文件WMASS.OUT
Close #i

'-------------------------------------------------------------------------------------------高亮最值
'Sheets("d_M").Cells.EntireColumn.AutoFit
'
'Num_All = Sheets("d_M").range("a65536").End(xlUp)
'Debug.Print "总楼层="; Num_All
'
''Dim ii As Integer
'Dim i_RowID As Integer
'Dim i_Rng As range
'
''---------------------------------------------------------刚度比
'For ii = 2 To 3
'Dim R As range
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 1, ii))
'Call maxormin(R, "min", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next


'-------------------------------------------------------------------------------------------读取最小刚度比
Sheets("g_M").Cells(22, 5).Formula = "=MIN(d_M!B" & Num_Base + 3 & ":B" & Num_all + 1 & ")"
Sheets("g_M").Cells(22, 7).Formula = "=MIN(d_M!C" & Num_Base + 3 & ":C" & Num_all + 1 & ")"

'-------------------------------------------------------------------------------------------对刚度比进行修正
For ii = 2 To 3
    Sheets("d_M").Cells(Num_Base + 3, ii).Interior.ColorIndex = 7
        For jj = 4 To Num_all + 1
            If Sheets("d_M").Cells(jj, 60).Value / Sheets("d_M").Cells(jj + 1, 60).Value > 1.5 Then
            Sheets("d_M").Cells(jj, ii).Interior.ColorIndex = 7
        End If
    Next
Next
'
'If i2 = 6 Then
'Num_all = Sheets("d_M").range("a65536").End(xlUp)
'For ii = 4 To 5
'    For jj = 3 To Num_all + 2
'    Sheets("d_M").Cells(jj, ii) = Sheets("d_M").Cells(jj, ii) * Sheets("d_M").Cells(jj, 60)
'    Next
'Next
'End If

Debug.Print "耗费时间: " & Timer - sngStart

End Sub
