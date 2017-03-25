Attribute VB_Name = "MBuilding_抗剪承载力"


'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/03/19
'1.添加判断代码，解决没有抗剪承载力报错问题。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15
'1.修改高亮代码，换为条件格式

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29
'1.增加高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/30

'更新内容：
'1.修正承载力的开始行，应为Num_Base + 3

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/18

'更新内容：
'1.


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****        MBuilding_楼层抗剪承载力突变验算.TXT部分代码                  ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************



Sub OUTReader_MBuilding_抗剪承载力(Path As String)

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
Dim Keyword_VRate1, Keyword_VRate2   As String
'赋值
Keyword_Rate1 = "RS_0"
Keyword_Rate2 = "RS_90"



'==========================================================================================定义首字符变量


'刚度比
Dim FirstString_kRate1 As String
Dim FirstString_kRate2 As String

'=============================================================================================================================生成文件读取路径

'指定文件名为wmass.out
Filename = Dir(Path & "\*_楼层抗剪承载力突变验算.txt")

'Dim Fliename1, Filename2 As String
'Filename1 = Dir(Path & "\*_结构总信息.txt")
'Debug.Print "AA" & Filename1
'Filename2 = CheckRegExpfromString(Fliename, "\w+_")


Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件


If Dir(Path & "\" & "*_楼层抗剪承载力突变验算.txt") <> "" Then  '------------------------------------------------------------------------------------------------------------------------------添加

Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件：楼层抗剪承载力突变验算.TXT"
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
       

    '-------------------------------------------------------------------------------------------读取抗剪承载力
    
        
    If FirstString_kRate1 = Keyword_Rate1 Then
        'Debug.Print "aaaaaaaaaaaaaa"
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

                Sheets("d_M").Cells(j, 46) = StringfromStringforReg(data, "\S+", 4)
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

                '逐一写入X向刚度及刚度比
                Sheets("d_M").Cells(j, 47) = StringfromStringforReg(data, "\S+", 4)
            End If
            If CheckRegExpfromString(data, "--") = True Then
                Exit Do
            End If
            
            ii = ii + 1
                       
        Loop
        Num_all = ii
        'Debug.Print Num_All

    End If
    
          
Loop

Sheets("d_M").Cells(ii + 2, 46) = 1
Sheets("d_M").Cells(ii + 2, 47) = 1

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


''---------------------------------------------------------承载力比
'For ii = 46 To 47
'Dim R As range
'Set R = Worksheets("d_M").range(Cells(3, ii), Cells(Num_All + 1, ii))
'Call maxormin(R, "min", "d_M!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next


'-------------------------------------------------------------------------------------------读取最大楼层抗剪承载力比
Sheets("g_M").Cells(23, 5).Formula = "=MIN(d_M!AT" & Num_Base + 3 & ":AT" & Num_all + 1 & ")"
Sheets("g_M").Cells(23, 7).Formula = "=MIN(d_M!AU" & Num_Base + 3 & ":AU" & Num_all + 1 & ")"

Debug.Print "耗费时间: " & Timer - sngStart

End If

End Sub

