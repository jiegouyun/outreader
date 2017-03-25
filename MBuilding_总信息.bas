Attribute VB_Name = "MBuilding_总信息"
Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/24
'1.添加判断代码,解决楼层没有质量是出错问题。

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/10
'1.添加代码,解决模型建立地下室的数据读取问题。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15
'1.添加读取层质量和层高
'2.修正高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29
'1.增加高亮代码

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/08/01

'更新内容：
'1.修改刚重比读入代码；


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/30

'更新内容：
'1.修改Num_All读取代码；
'2.写入路径、程序、计算日期、楼层自由度、周期折减系数等信息；


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/07/18

'更新内容：
'1.


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                  MBuilding_结构总信息.TXT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************



Sub OUTReader_MBuilding_总信息(Path As String)

'计算运行时间
Dim sngStart As Single
sngStart = Timer


'==========================================================================================定义主要辅助变量

'定义文件路径、文件名、文件完整路径、字符变量
Dim Filename  As String, filepath  As String, inputstring  As String

'定义data为读入行的字符串
Dim data As String


'定义循环变量
Dim i As Integer, j As Integer, NN As Integer

Dim i_m As Integer, i_k1 As Integer, i_k2 As Integer, i_w As Integer

'i_k1、i_k2分别为两种刚度比的写入行数记录，第3行为第1层，前两行为标题行
i_k1 = 3
i_k2 = 3


'文本当前行数
Dim n As Integer

Dim Mass, kWeight


'==========================================================================================定义关键词变量

'质量
Dim Keyword_Mass1 As String, Keyword_Mass2 As String, Keyword_Mass3 As String, Keyword_Mass4 As String
'赋值
Keyword_Mass1 = "活载产生的总质量(t)"
Keyword_Mass2 = "恒载产生的总质量(t)"
Keyword_Mass3 = "附加总质量 (t):  " '------------------------MB里没这项
Keyword_Mass4 = "结构的总质量(t)"

'单位面积质量
Dim Keyword_MassAve As String
'赋值
Keyword_MassAve = "层质量比验算结果" '----------------------MB里没有直接输出单位面积质量，只有层的总质量

'风荷载
Dim Keyword_Wind As String
'赋值
Keyword_Wind = "规范方法" '---------------------------------MB里结构倾覆力矩有两种输出结果：规范方法和结构力学方法，此处采用规范方法


'倾覆力矩
Dim Keyword_Over As String
'赋值
Keyword_Over = "抗倾覆弯矩ROTM"

'楼层属性'--------------------------------?
Dim Keyword_Flo As String
'赋值
Keyword_Flo = "楼层属性"

'刚重比
Dim Keyword_kWeight0, Keyword_kWeight1, Keyword_kWeight2 As String '-----------MB怎么两组刚重比呢？
'赋值
Keyword_kWeight0 = "EJd"
Keyword_kWeight1 = " RS_0"
Keyword_kWeight2 = "RS_90"

'赋值
Dim Keyword_FloShear  As String
Keyword_FloShear = "Ratio_Bu"

'周期折减系数
Dim Keyword_TD, Keyword_TA As String
Keyword_TD = "周期折减系数"

'计算振型数
Keyword_TA = "计算振型数"


'==========================================================================================定义首字符变量

'质量
Dim FirstString_Mass As String
Dim FirstString_Mass_2 As String

'单位面积质量
Dim FirstString_MassAve As String


'风荷载
Dim FirstString_Wind As String

'倾覆力矩
Dim FirstString_Over As String

'楼层属性
Dim FirstString_Flo As String

'刚重比
Dim FirstString_kWeight As String


'周期折减系数
Dim FirstString_TD As String

'计算振型数
Dim FirstString_TA As String



'=============================================================================================================================生成文件读取路径

'指定文件名为wmass.out
Filename = Dir(Path & "\*_结构总信息.txt")

Dim i_ As Integer: i = FreeFile

'生成完整文件路径
filepath = Path & "\" & Filename
Debug.Print Path
Debug.Print filepath

'打开结果文件
Open (filepath) For Input Access Read As #i


'=============================================================================================================================逐行读取文本

Debug.Print "开始遍历结果文件wmass.out"
Debug.Print "读取相关指标"
Debug.Print "……"


Do While Not EOF(1)

    Line Input #i, inputstring '读文本文件一行
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符
    
    '质量
    FirstString_Mass = Mid(data, 3, 11)
    FirstString_Mass_2 = Mid(data, 3, 9)
    '单位面积质量
    FirstString_MassAve = Mid(data, 6, 8)
    '风荷载
    FirstString_Wind = Mid(data, 52, 4)
    '抗倾覆
    FirstString_Over = Mid(data, 13, 9)
    '楼层属性'-----------------------------------?
    FirstString_Flo = Mid(data, 30, 4)
    '刚重比
    'FirstString_kWeight = Mid(data, 6, 15) '-----------
    
    '周期折减系数
    FirstString_TD = Mid(data, 3, 6)
    '计算振型数
    FirstString_TA = Mid(data, 3, 5)
    
    
    If Mid(data, 64, 4) = "计算日期" Then
        'CalTime = Mid(data, 57, 11)
        'Debug.Print "计算时间:" & CalTime
        Sheets("g_M").Cells(4, 7) = Mid(data, 70, 8)
    End If
    

    '-------------------------------------------------------------------------------------------读取结构体系及层数信息
    'Dim StrType As Integer
    'StrType = 1
    'If Mid(data, 3, 5) = "结构体系:" Then
        'If CheckRegExpfromString(data, "结构体系:\s+\w*剪力墙\w*") Then
            'StrType = 2
            '刚度比重新赋值
            'Keyword_Rate = "Ratx2="
            'Debug.Print data
        'End If
    'End If

    If Mid(data, 3, 6) = "地下室层数:" Then
        Num_Base = StringfromStringforReg(data, "\S+", 2)
        Debug.Print "地下室层数:"; Num_Base
    End If

    Dim Num_Change As Integer
    If Mid(data, 6, 8) = "转换层所在楼层：" Then
        Num_Change = StringfromStringforReg(data, "\S+", 2)
        Debug.Print "转换层所在层号："; Num_Change
    End If


    If FirstString_TA = Keyword_TA Then
        Sheets("g_M").Cells(38, 7) = extractNumberFromString(data, 1)
    End If
    
    If FirstString_TD = Keyword_TD Then
        Sheets("g_M").Cells(5, 7) = extractNumberFromString(data, 1)
    End If
    
    
'        '-------------------------------------------------------------------------------------------读取层质量------------------------------------------添加
'
'    If Mid(data, 6, 5) = "各层的质量" Then
'        Line Input #i, data
'        Line Input #i, data
'        Line Input #i, data
'        Line Input #i, data
'        Do While Not EOF(1)
'            Line Input #i, data
'            '如果接连两个数，认为是数据行
'
'                '结果文件中数据对应层号从大至小，统一为从小到大排列
'                'j为读取行数据写入表格的行数，跳过两行标题行
'                j = extractNumberFromString(data, 1) + 2
'
'                '逐一写入质量
'                Sheets("d_M").Cells(j, 54) = extractNumberFromString(data, 5) + extractNumberFromString(data, 6)
'                Sheets("d_M").Cells(j, 54) = Round(Sheets("d_M").Cells(j, 54), 1)
'                'Sheets("d_M").Cells(j, 54) = StringfromStringforReg(data, "\S+", 6) & "+" & StringfromStringforReg(data, "\S+", 7)
'            If CheckRegExpfromString(data, "--") = True Then
'                Exit Do
'            End If
'        Loop

'        Debug.Print "读取风荷载信息"
'        Debug.Print "用时: " & Timer - sngStart
'        Debug.Print "……"
'
'    End If

    '索引字符符合“活荷载”关键词
    If FirstString_Mass = Keyword_Mass1 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_M").Cells(6, 5) = extractNumberFromString(data, 1)

        Debug.Print "读取活荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“恒荷载”关键词
    If FirstString_Mass = Keyword_Mass2 Then

         '提取质量数据，并写入工作表相应位置
        Sheets("g_M").Cells(7, 5) = extractNumberFromString(data, 1)


        Debug.Print "读取恒荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“附加荷载”关键词
    If FirstString_Mass = Keyword_Mass3 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_M").Cells(9, 5) = extractNumberFromString(data, 1)

        Debug.Print "读取附加荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“总质量”关键词
    If FirstString_Mass_2 = Keyword_Mass4 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_M").Cells(7, 7) = extractNumberFromString(data, 1)

        Debug.Print "读取结构总质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    
            '-------------------------------------------------------------------------------------------读取层高------------------------------------------添加

    If Mid(data, 6, 6) = "各层构件数量" Then
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then
                If CheckRegExpfromString(data, "B\S\F") = False Then

                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                    Debug.Print j & "11111111111111111111111111111111"
                End If
                Sheets("d_M").Cells(j, 60) = StringfromStringforReg(data, "\S+", 8)
            End If
            If CheckRegExpfromString(data, "--") = True Then
                Exit Do
            End If
        Loop

        Debug.Print "读取层高"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If


    
    '-------------------------------------------------------------------------------------------读取风荷载信息

    If FirstString_Wind = Keyword_Wind Then
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then
                If CheckRegExpfromString(data, "B\S\F") = False Then

                    '结果文件中数据对应层号从大至小，统一为从小到大排列
                    'j为读取行数据写入表格的行数，跳过两行标题行
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If

                '逐一写入层号、剪力X、倾覆弯矩X、剪力Y、倾覆弯矩Y
                Sheets("d_M").Cells(j, 1) = j - 2
                Sheets("d_M").Cells(j, 6) = StringfromStringforReg(data, "\S+", 4)
                Sheets("d_M").Cells(j, 7) = StringfromStringforReg(data, "\S+", 5)
                Sheets("d_M").Cells(j, 8) = StringfromStringforReg(data, "\S+", 7)
                Sheets("d_M").Cells(j, 9) = StringfromStringforReg(data, "\S+", 8)

            End If
            If CheckRegExpfromString(data, "--") = True Then
                Exit Do
            End If
        Loop

        Debug.Print "读取风荷载信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If


    '-------------------------------------------------------------------------------------------读取抗倾覆信息

    If FirstString_Over = Keyword_Over Then

        '按表格格式，从26行开始写入数据
        j = 48
    
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){3}") = True Then

                '逐一写入抗倾覆力矩、倾覆力矩、比值、零应力区
                Sheets("g_M").Cells(j, 4) = extractNumberFromString(data, 1)
                Sheets("g_M").Cells(j, 5) = extractNumberFromString(data, 2)
                Sheets("g_M").Cells(j, 6) = extractNumberFromString(data, 3)
                Sheets("g_M").Cells(j, 7) = "-"

                '换行
                j = j + 1

            End If

            '遇到分隔符“==”则退出小循环
            If CheckRegExpfromString(data, "==") = True Then

                '输出提示语
                Debug.Print "读取抗倾覆信息"
                Debug.Print "用时: " & Timer - sngStart
                Debug.Print "……"

                Exit Do

            End If

        Loop

    End If

    '-------------------------------------------------------------------------------------------读取刚重比
    If CheckRegExpfromString(data, "\s刚重比") Then
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data
                '刚重比
            FirstString_kWeight = Mid(data, 6, 5) '-----------
            
             'X向刚重比判断
            If FirstString_kWeight = Keyword_kWeight1 Then
'           Debug.Print FirstString_kWeight
'           Debug.Print data

            Sheets("g_M").Cells(20, 5) = extractNumberFromString(data, 4)
            End If


            'Y向刚重比判断
             If FirstString_kWeight = Keyword_kWeight2 Then
'            Debug.Print FirstString_kWeight
'           Debug.Print data

            Sheets("g_M").Cells(21, 5) = extractNumberFromString(data, 4)
            End If
            
            If CheckRegExpfromString(data, "==") = True Then
                Exit Do
            End If
        Loop
    End If
    
    
    '-------------------------------------------------------------------------------------------读取单位面积质量
    NN = 0
    If FirstString_MassAve = Keyword_MassAve Then
        Do While Not EOF(1)
            Line Input #i, data
                '刚重比
            FirstString_kWeight = Mid(data, 6, 15) '-----------
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then
                If CheckRegExpfromString(data, "B\S\F") = False Then
                    j = extractNumberFromString(data, 1) + 2 + Num_Base
                Else
                    j = Num_Base - CInt(Mid(StringfromStringforReg(data, "B\S\F", 1), 2, 1)) + 1 + 2
                End If

                '逐一写入层号、单位面积质量、质量比
                Dim aa As String
                aa = StringfromStringforReg(data, "\S+", 3)
                If aa = "-" Then
                    Sheets("d_M").Cells(j, 54) = 1
                    Sheets("d_M").Cells(j, 55) = 1
                Else
                    Sheets("d_M").Cells(j, 54) = aa / 1000
                    Sheets("d_M").Cells(j, 55) = StringfromStringforReg(data, "\S+", 4)
                End If
'                Sheets("d_M").Cells(j, 55) = StringfromStringforReg(data, "\S+", 4)
                NN = NN + 1

            End If
            'If CheckRegExpfromString(data, "==") = True Then
                '记录总层数
                'Num_All = j - 2
                'Debug.Print Num_All
                'Exit Do
            'End If
        Loop

    End If
'
    
Loop

Num_all = NN
Debug.Print Num_all

'关闭结果文件

'-------------------------------------------------------------------------------------------高亮最值
'Sheets("d_M").Cells.EntireColumn.AutoFit
'
'Num_All = Sheets("d_M").range("a65536").End(xlUp)
'Debug.Print "总楼层="; Num_All
'
'Dim ii As Integer
'Dim i_RowID As Integer
'Dim i_Rng As range
'
'
''---------------------------------------------------------质量比
'For ii = 55 To 55
'Dim R As range
'Set R = Worksheets("d_M").range(Cells(4, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_M!R4C" & CStr(ii) & ":R" & CStr(Num_All + 2) & "C" & CStr(ii))
'Next


Close #i
'-------------------------------------------------------------------------------------------判断剪重比情况
Select Case Sheets("g_M").Cells(20, 5).Value
    Case Is < 1.4: Sheets("g_M").Cells(20, 7) = "稳定不足,考虑二阶"
    Case 1.4 To 2.7: Sheets("g_M").Cells(20, 7) = "满足稳定,考虑二阶"
    Case Is > 2.7: Sheets("g_M").Cells(20, 7) = "满足稳定,不计二阶"
End Select

Select Case Sheets("g_M").Cells(21, 5).Value
    Case Is < 1.4: Sheets("g_M").Cells(21, 7) = "稳定不足,考虑二阶"
    Case 1.4 To 2.7: Sheets("g_M").Cells(21, 7) = "满足稳定,考虑二阶"
    Case Is > 2.7: Sheets("g_M").Cells(21, 7) = "满足稳定,不计二阶"
End Select

'-------------------------------------------------------------------------------------------读取首层风荷载下的剪力和弯矩
'X向剪力
Sheets("g_M").Cells(42, 4).Formula = "=d_M!F" & Num_Base + 3
'X向弯矩
Sheets("g_M").Cells(42, 6).Formula = "=d_M!G" & Num_Base + 3
'Y向剪力
Sheets("g_M").Cells(43, 4).Formula = "=d_M!H" & Num_Base + 3
'Y向弯矩
Sheets("g_M").Cells(43, 6).Formula = "=d_M!I" & Num_Base + 3

'-------------------------------------------------------------------------------------------写入数据路径和计算程序名称
Sheets("g_M").Cells(3, 4) = OUTReader_Main.TextBox_Path.Text
Sheets("g_M").Cells(4, 4) = "Midas Building"

'-------------------------------------------------------------------------------------------将单位面积质量改为楼层质量
Sheets("d_M").Cells(1, 54) = "楼层质量分布"
Sheets("d_M").Cells(2, 54) = "楼层质量"
Sheets("d_M").Cells(Num_Base + 3, 55) = 1



Sheets("d_M").Cells(1, 60) = "层高"
Sheets("d_M").Cells(2, 60) = "m"


Debug.Print "耗费时间: " & Timer - sngStart

End Sub

