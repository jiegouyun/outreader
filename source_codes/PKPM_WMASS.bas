Attribute VB_Name = "PKPM_WMASS"
Public Num_all, Num_Base As Integer

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                            PKPM_WMASS.OUT部分代码                    ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/11/21
'更新内容:
'1.匹配PKPM v2.2 版本。更新内容包括：1）地下室层数 2)风荷载提取 3) 层高


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/07/057
'更新内容:
'1.修正PKPM刚度比修正时起始行数错误。（YJK正确）

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/05/05
'更新内容:
'1.修正结构类型的判断，增加“框剪”的判断


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/03/18
'更新内容:
'1.修正刚度修正模块，对刚度比进行修正时，从第3行开始


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/01/09
'更新内容:
'1.隐去高亮代码；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15
'更新内容:
'1.添加读取层高
'2.修改读取层质量(1.0D+1.0L)和质量比
'3.修正刚度比、承载力比、质量比的高亮代码
'4.修正首层风荷载下剪力弯矩公式：+4改为+3，地震的没找到在哪里啊？
'5.添加过程对刚度及刚度比进行修正，以和midas比较；按钮放在“模型对比”功能项里

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/10/29
'更新内容:
'1.修正结构类型的判断，加入“筒”的判断

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/29
'更新内容:
'1.添加高亮代码
'2.刚度比和承载力比取有效数字

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'更新内容:
'1.修正刚性楼板假定读取bug
'2.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/29
'更新内容:
'1.增加计算日期的读取

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/19
'更新内容:
'1.修正首层“楼层抗剪承载力”为0的bug
'2.首层风荷载下的剪力和弯矩写入general

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/16
'更新内容:
'1.增加读取振型个数

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/12
'更新内容:
'1.适配新版general修改输入信息的位置及内容


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/8

'更新内容：
'1.修改刚度比及抗剪承载力比的公式；

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/6 20:44

'更新内容：
'1.增加总层数和地下室层数的全局变量定义

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/3 20:45

'更新内容：
'1.更正刚重比写入位置

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/27 19:37

'更新内容：
'1.修改每层单位质量列的位置

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/5/21 21:15

'更新内容：
'1.删除无用的多余变量Row_over、Row_wind、Row_FloShear、Row_Flo
'2.Num_All赋值为全楼楼层数（在读取单位面积质量结束时存储），Num_All需在MainProgram里定义为全局变量
'3.添加了最大刚度比、最大楼层承载力比
'4.添加单位面积质量、质量比读取

'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2013/5/17 18:01

'////////////////////////////////////////////////////////////////////////////////////////////



Sub OUTReader_PKPM_WMASS(Path As String)

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

Dim i_m, i_k1, i_k2, i_w As Integer

'i_k1、i_k2分别为两种刚度比的写入行数记录，第3行为第1层，前两行为标题行
i_k1 = 3
i_k2 = 3


'文本当前行数
Dim n As Integer

Dim Mass, kWeight


'==========================================================================================定义关键词变量

'质量
Dim Keyword_Mass1, Keyword_Mass2, Keyword_Mass3, Keyword_Mass4 As String
'赋值
Keyword_Mass1 = "活载产生的总质量 (t)"
Keyword_Mass2 = "恒载产生的总质量 (t)"
Keyword_Mass3 = "附加总质量 (t):  "
Keyword_Mass4 = "结构的总质量 (t): "

'单位面积质量
Dim Keyword_MassAve As String
'赋值
Keyword_MassAve = "层号   塔号    单位面积质量"

'风荷载
Dim Keyword_Wind1 As String
Dim Keyword_Wind2 As String
'赋值
Keyword_Wind1 = "层号  塔号   风荷载X" '无横向和扭转
Keyword_Wind2 = "层号  塔号 风作用方向"

'刚度比
Dim Keyword_Rate As String
'赋值
Keyword_Rate = "Ratx1="

'刚度
Dim Keyword_K As String
'赋值
Keyword_K = "RJX3 ="

'倾覆力矩
Dim Keyword_Over As String
'赋值
Keyword_Over = "抗倾覆力矩Mr"

'楼层属性
Dim Keyword_Flo As String
'赋值
Keyword_Flo = "楼层属性"

'刚重比
Dim Keyword_kWeight1, Keyword_kWeight2 As String
'赋值
Keyword_kWeight1 = "X向刚重比"
Keyword_kWeight2 = "Y向刚重比"

'楼层承载力比
Dim Keyword_FloShear As String
'赋值
Keyword_FloShear = "Ratio_Bu"


'==========================================================================================定义首字符变量

'质量
Dim FirstString_Mass As String

'单位面积质量
Dim FirstString_MassAve As String

'刚度比
Dim FirstString_kRate As String

'刚度比
Dim FirstString_K As String

'风荷载
Dim FirstString_Wind1 As String
Dim FirstString_Wind2 As String

'倾覆力矩
Dim FirstString_Over As String

'楼层属性
Dim FirstString_Flo As String

'刚重比
Dim FirstString_kWeight As String

'楼层承载力
Dim Firststring_FloShear As String


'=============================================================================================================================生成文件读取路径

'指定文件名为wmass.out
Filename = "WMASS.OUT"

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

    '记录行数
    n = n + 1
    
    '将读取的一行字符串赋值与data变量
    data = inputstring
    
    '-------------------------------------------------------------------------------------------定义各指标的判别字符
    
    '质量
    FirstString_Mass = Mid(data, 6, 12)
    '单位面积质量
    FirstString_MassAve = Mid(data, 2, 17)
    '风荷载
    FirstString_Wind1 = Mid(data, 2, 13) '无横向扭转
    FirstString_Wind2 = Mid(data, 2, 12)
    '刚度比
    FirstString_kRate = Mid(data, 3, 6)
    '刚度
    FirstString_K = Mid(data, 3, 6)
    '抗倾覆
    FirstString_Over = Mid(data, 14, 7)
    '楼层属性
    FirstString_Flo = Mid(data, 30, 4)
    '刚重比
    FirstString_kWeight = Mid(data, 2, 5)
    '楼层剪力
    Firststring_FloShear = Mid(data, 11, 8)
    

    '-------------------------------------------------------------------------------------------读取结构体系及层数信息及模型信息
    Dim CalTime As String
    If Mid(data, 54, 2) = "日期" Then
        CalTime = Mid(data, 57, 11)
        Debug.Print "计算时间:" & CalTime
        Sheets("g_P").Cells(4, 7) = CalTime
    End If
    
    Dim StrType As Integer
    StrType = 1
    If Mid(data, 6, 5) = "结构类别:" Then
        If CheckRegExpfromString(data, "结构类别:\s+.*剪.*") Or CheckRegExpfromString(data, "结构类别:\s+.*筒.*") Then
            StrType = 2
            '刚度比重新赋值
            Keyword_Rate = "Ratx2="
            Debug.Print data
        End If
    End If

    If Mid(data, 6, 6) = "地下室层数:" Then
'        Num_Base = StringfromStringforReg(data, "\S+", 3) 'PKPM 2010版
        Num_Base = StringfromStringforReg(data, "\S+", 3)
        Debug.Print "地下室层数:"; Num_Base
    End If

    Dim Num_Change As Integer
    If Mid(data, 6, 8) = "转换层所在层号：" Then
        Num_Change = StringfromStringforReg(data, "\S+", 3)
        Debug.Print "转换层所在层号："; Num_Change
    End If


    Sheets("g_P").Cells(4, 4) = "SATWE 中文版"

    If Mid(data, 6, 15) = "是否对全楼强制采用刚性楼板假定" Then
        If Mid(data, 29, 1) = "是" Then Sheets("g_P").Cells(5, 5) = "刚性楼板假定"
        If Mid(data, 29, 1) = "否" Then Sheets("g_P").Cells(5, 5) = "非刚性楼板假定"
    End If

    If Mid(data, 6, 6) = "周期折减系数" Then
        Sheets("g_P").Cells(5, 7) = extractNumberFromString(data, 1)
    End If


    If Mid(data, 6, 5) = "计算振型数" Then
        Sheets("g_P").Cells(38, 7) = extractNumberFromString(data, 1)
    End If
    
    Sheets("g_P").Cells(38, 7).NumberFormatLocal = "0"
    
            '-------------------------------------------------------------------------------------------读取层质量
    Dim ks As String

    If Mid(data, 25, 5) = "各层的质量" Then
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then

                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2

                '逐一写入层质量
                Sheets("d_P").Cells(j, 54) = extractNumberFromString(data, 6) + extractNumberFromString(data, 7)
                Sheets("d_P").Cells(j, 54) = Round(Sheets("d_P").Cells(j, 54), 1)
                'Sheets("d_P").Cells(j, 54) = StringfromStringforReg(data, "\S+", 6) & "+" & StringfromStringforReg(data, "\S+", 7)
  

            End If
            If j = 3 Then
                Exit Do
            End If
        Loop

        Debug.Print "读取层质量信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    

    '-------------------------------------------------------------------------------------------读取质量


'    Select Case FirstString_Mass
'
'        Case Keyword_Mass1
'
'            '提取质量数据，并写入工作表相应位置
'            Sheets("g_P").Cells(7, 5) = extractNumberFromString(data, 1)
'
'            Debug.Print "读取活荷载质量"
'            Debug.Print "用时: " & Timer - sngStart
'            Debug.Print "……"
'
'        Case Keyword_Mass2
'
'            '提取质量数据，并写入工作表相应位置
'            Sheets("g_P").Cells(8, 5) = extractNumberFromString(data, 1)
'
'            Debug.Print "读取恒荷载质量"
'            Debug.Print "用时: " & Timer - sngStart
'            Debug.Print "……"
'
'        Case Keyword_Mass3
'
'            '提取质量数据，并写入工作表相应位置
'            Sheets("g_P").Cells(9, 5) = extractNumberFromString(data, 1)
'
'            Debug.Print "读取附加荷载质量"
'            Debug.Print "用时: " & Timer - sngStart
'            Debug.Print "……"
'
'        Case Keyword_Mass4
'
'            '提取质量数据，并写入工作表相应位置
'            Sheets("g_P").Cells(11, 5) = extractNumberFromString(data, 1)
'
'            Debug.Print "读取结构总质量"
'            Debug.Print "用时: " & Timer - sngStart
'            Debug.Print "……"
'
'    End Select



    '索引字符符合“活荷载”关键词
    If FirstString_Mass = Keyword_Mass1 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_P").Cells(6, 5) = extractNumberFromString(data, 1)

        Debug.Print "读取活荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“恒荷载”关键词
    If FirstString_Mass = Keyword_Mass2 Then

         '提取质量数据，并写入工作表相应位置
        Sheets("g_P").Cells(7, 5) = extractNumberFromString(data, 1)


        Debug.Print "读取恒荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“附加荷载”关键词
    If FirstString_Mass = Keyword_Mass3 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_P").Cells(6, 7) = extractNumberFromString(data, 1)

        Debug.Print "读取附加荷载质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If

    '索引字符符合“总质量”关键词
    If FirstString_Mass = Keyword_Mass4 Then

        '提取质量数据，并写入工作表相应位置
        Sheets("g_P").Cells(7, 7) = extractNumberFromString(data, 1)

        Debug.Print "读取结构总质量"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    
                '-------------------------------------------------------------------------------------------读取层高

    If Mid(data, 20, 6) = "各层构件数量" Then

        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Line Input #i, data
        Do While Not EOF(1)
            Line Input #i, data

                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2
                '逐一写入层高
'                Sheets("d_P").Cells(j, 60) = StringfromStringforReg(data, "\S+", 10) 'PKPM v2010
                Sheets("d_P").Cells(j, 60) = StringfromStringforReg(data, "\S+", 14)
            If CheckRegExpfromString(data, "\*") = True Then
                Exit Do
            End If
        Loop

        Debug.Print "读取层质量信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If



    '-------------------------------------------------------------------------------------------读取单位面积质量

    If FirstString_MassAve = Keyword_MassAve Then
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then

                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2

                '逐一写入层号、单位面积质量、质量比
                Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                'Sheets("d_P").Cells(j, 54) = StringfromStringforReg(data, "\S+", 3)'-------------------------------
                Sheets("d_P").Cells(j, 55) = StringfromStringforReg(data, "\S+", 4)

            End If
            If CheckRegExpfromString(data, "==") = True Then
                '记录总层数
                Num_all = j - 2
                Exit Do
            End If
        Loop

        Debug.Print "读取风荷载信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If


    
    '-------------------------------------------------------------------------------------------读取风荷载信息

'    无横风向和扭转
    If FirstString_Wind1 = Keyword_Wind1 Then
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then

                '结果文件中数据对应层号从大至小，统一为从小到大排列
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2

                '逐一写入层号、剪力X、倾覆弯矩X、剪力Y、倾覆弯矩Y
                Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                Sheets("d_P").Cells(j, 6) = StringfromStringforReg(data, "\S+", 4)
                Sheets("d_P").Cells(j, 7) = StringfromStringforReg(data, "\S+", 5)
                Sheets("d_P").Cells(j, 8) = StringfromStringforReg(data, "\S+", 7)
                Sheets("d_P").Cells(j, 9) = StringfromStringforReg(data, "\S+", 8)

            End If
            If CheckRegExpfromString(data, "==") = True Then
                Exit Do
            End If
        Loop

        Debug.Print "读取风荷载信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    
'   有横风向和扭转
    If FirstString_Wind2 = Keyword_Wind2 Then
        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){2}") = True Then

                '结果文件中数据对应层号从大至小，统一为从小到大排列
                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2

                '逐一写入层号、剪力X、倾覆弯矩X、剪力Y、倾覆弯矩Y
                Sheets("d_P").Cells(j, 1) = StringfromStringforReg(data, "\S+", 1)
                Sheets("d_P").Cells(j, 6) = StringfromStringforReg(data, "\S+", 5)
                Sheets("d_P").Cells(j, 7) = StringfromStringforReg(data, "\S+", 6)
                
                Line Input #i, data
                
                Sheets("d_P").Cells(j, 8) = StringfromStringforReg(data, "\S+", 3)
                Sheets("d_P").Cells(j, 9) = StringfromStringforReg(data, "\S+", 4)
                

            End If
            If CheckRegExpfromString(data, "==") = True Then
                Exit Do
            End If
        Loop

        Debug.Print "读取风荷载信息"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
    
    

    '-------------------------------------------------------------------------------------------读取刚度比
    
        
    If FirstString_kRate = Keyword_Rate Then
'       Debug.Print data
'       Debug.Print extractNumberFromString(data, 2)
        Sheets("d_P").Cells(i_k1, 2) = extractNumberFromString(data, 1)
        Sheets("d_P").Cells(i_k1, 2) = Round(Sheets("d_P").Cells(i_k1, 2), 4)
        Sheets("d_P").Cells(i_k1, 3) = extractNumberFromString(data, 2)
        Sheets("d_P").Cells(i_k1, 3) = Round(Sheets("d_P").Cells(i_k1, 3), 4)
        i_k1 = i_k1 + 1
    End If


    '-------------------------------------------------------------------------------------------读取刚度


    If FirstString_K = Keyword_K Then
'       Debug.Print data
'       Debug.Print extractNumberFromString(data, 2)
        Sheets("d_P").Cells(i_k2, 4) = extractNumberFromString(data, 1)
        Sheets("d_P").Cells(i_k2, 5) = extractNumberFromString(data, 2)
        i_k2 = i_k2 + 1
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
                Sheets("g_P").Cells(j, 4) = extractNumberFromString(data, 1)
                Sheets("g_P").Cells(j, 5) = extractNumberFromString(data, 2)
                Sheets("g_P").Cells(j, 6) = extractNumberFromString(data, 3)
                Sheets("g_P").Cells(j, 7) = extractNumberFromString(data, 4)

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

    'X向刚重比判断
    If FirstString_kWeight = Keyword_kWeight1 Then
'        Debug.Print FirstString_kWeight
'        Debug.Print data

        Sheets("g_P").Cells(20, 5) = extractNumberFromString(data, 1)

    End If


    'Y向刚重比判断
    If FirstString_kWeight = Keyword_kWeight2 Then
'        Debug.Print FirstString_kWeight
'        Debug.Print data

        Sheets("g_P").Cells(21, 5) = extractNumberFromString(data, 1)

    End If



    '-------------------------------------------------------------------------------------------读取楼层承载力验算

    If Firststring_FloShear = Keyword_FloShear Then

        Do While Not EOF(1)
            Line Input #i, data
            '如果接连两个数，认为是数据行
            If CheckRegExpfromString(data, "(\s*(\-?)(\d*)(\.?)(\d*)([E]?)([+]?)([-]?)(\d+)){3}") = True Then

                '结果文件中数据对应层号从大至小，统一为从小到大排列

                'j为读取行数据写入表格的行数，跳过两行标题行
                j = extractNumberFromString(data, 1) + 2

                '逐一写入Ratio_X、Ratio_Y
                Sheets("d_P").Cells(j, 46) = extractNumberFromString(data, 5)
                Sheets("d_P").Cells(j, 46) = Round(Sheets("d_P").Cells(j, 46), 2)
                Sheets("d_P").Cells(j, 47) = extractNumberFromString(data, 6)
                Sheets("d_P").Cells(j, 47) = Round(Sheets("d_P").Cells(j, 47), 2)

            End If
            If j = 3 Then
                Exit Do
            End If
        Loop

        Debug.Print "读取楼层承载力验算"
        Debug.Print "用时: " & Timer - sngStart
        Debug.Print "……"

    End If
'
    
Loop

'关闭结果文件WMASS.OUT
Close #i

'-------------------------------------------------------------------------------------------重新计算质量比
Sheets("d_P").Cells(1, 54) = "楼层质量分布"
Sheets("d_P").Cells(2, 54) = "楼层质量"
Sheets("d_P").Cells(3, 55).Value = 1
For ii = 4 To Num_all + 2
    Sheets("d_P").Cells(ii, 55).Value = Sheets("d_P").Cells(ii, 54).Value / Sheets("d_P").Cells(ii - 1, 54).Value
    Sheets("d_P").Cells(ii, 55).Value = Round(Sheets("d_P").Cells(ii, 55).Value, 2)
Next




'-------------------------------------------------------------------------------------------读取最小刚度比
'Sheets("g_P").Cells(11, 14).Formula = "=MIN(d_P!B:B)"
Sheets("g_P").Cells(22, 5).Formula = "=MIN(d_P!B" & Num_Base + 3 & ":B" & Num_all + 1 & ")"
'Sheets("g_P").Cells(11, 15).Formula = "=MIN(d_P!C:C)"
Sheets("g_P").Cells(22, 7).Formula = "=MIN(d_P!C" & Num_Base + 3 & ":C" & Num_all + 1 & ")"

'-------------------------------------------------------------------------------------------读取最大楼层抗剪承载力比
'Sheets("g_P").Cells(13, 14).Formula = "=Min(d_P!AT:AT)"
Sheets("g_P").Cells(23, 5).Formula = "=MIN(d_P!AT" & Num_Base + 3 & ":AT" & Num_all + 2 & ")"
'Sheets("g_P").Cells(13, 15).Formula = "=Min(d_P!AU:AU)"
Sheets("g_P").Cells(23, 7).Formula = "=MIN(d_P!AU" & Num_Base + 3 & ":AU" & Num_all + 2 & ")"

'-------------------------------------------------------------------------------------------读取首层风荷载下的剪力和弯矩
'X向剪力
Sheets("g_P").Cells(42, 4).Formula = "=d_P!F" & Num_Base + 3
'X向弯矩
Sheets("g_P").Cells(42, 6).Formula = "=d_P!G" & Num_Base + 3
'Y向剪力
Sheets("g_P").Cells(43, 4).Formula = "=d_P!H" & Num_Base + 3
'Y向弯矩
Sheets("g_P").Cells(43, 6).Formula = "=d_P!I" & Num_Base + 3


'-------------------------------------------------------------------------------------------刚重比结果判定
Select Case Sheets("g_P").Cells(20, 5).Value
    Case Is < 1.4: Sheets("g_P").Cells(20, 7) = "稳定不足,考虑二阶"
    Case 1.4 To 2.7: Sheets("g_P").Cells(20, 7) = "满足稳定,考虑二阶"
    Case Is > 2.7: Sheets("g_P").Cells(20, 7) = "满足稳定,不计二阶"
End Select

Select Case Sheets("g_P").Cells(21, 5).Value
    Case Is < 1.4: Sheets("g_P").Cells(21, 7) = "稳定不足,考虑二阶"
    Case 1.4 To 2.7: Sheets("g_P").Cells(21, 7) = "满足稳定,考虑二阶"
    Case Is > 2.7: Sheets("g_P").Cells(21, 7) = "满足稳定,不计二阶"
End Select



Sheets("d_P").Cells(1, 60) = "层高"
Sheets("d_P").Cells(2, 60) = "m"


'-------------------------------------------------------------------------------------------高亮最值
'Sheets("d_P").Cells.EntireColumn.AutoFit
''
'Num_All = Sheets("d_P").range("a65536").End(xlUp)
'Debug.Print "总楼层="; Num_All
'
'Dim i_RowID As Integer
'Dim i_Rng As range
'
''---------------------------------------------------------刚度比
'For ii = 2 To 3
'Dim R As range
'Set R = Worksheets("d_P").range(Cells(3, ii), Cells(Num_All + 1, ii))
'Call maxormin(R, "min", "d_P!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
''---------------------------------------------------------承载力比
'For ii = 46 To 47
'Set R = Worksheets("d_P").range(Cells(3, ii), Cells(Num_All + 1, ii))
'Call maxormin(R, "min", "d_P!R3C" & CStr(ii) & ":R" & CStr(Num_All + 1) & "C" & CStr(ii))
'Next
'
''---------------------------------------------------------质量比
'ii = 55
'Set R = Worksheets("d_P").range(Cells(4, ii), Cells(Num_All + 2, ii))
'Call maxormin(R, "max", "d_P!R4C" & CStr(ii) & ":R" & CStr(Num_All + 2) & "C" & CStr(ii))
'
''
''Sheets("d_P").Cells.EntireColumn.AutoFit

Debug.Print "耗费时间: " & Timer - sngStart

End Sub



'------------------------------------------------------------------------------------------------------------------------添加过程对刚度比进行修正
Sub modi_stiff()
'-------------------------------------------------------------------------------------------对刚度进行层高修正
Num_all = Sheets("d_P").range("a65536").End(xlUp)
For ii = 4 To 5
    For jj = 3 To Num_all + 2
    Sheets("d_P").Cells(jj, ii) = Sheets("d_P").Cells(jj, ii) * Sheets("d_P").Cells(jj, 60)
    Next
Next

'-------------------------------------------------------------------------------------------对刚度比进行修正
For ii = 2 To 3
    'Sheets("d_P").Cells(Num_Base+3, ii).Value = Sheets("d_P").Cells(3, ii).Value * 1.5 '-------------------------默认首层嵌固
    Sheets("d_P").Cells(Num_Base + 3, ii).Interior.ColorIndex = 7
        For jj = 3 To Num_all + 1
            If Sheets("d_P").Cells(jj, 60).Value / Sheets("d_P").Cells(jj + 1, 60).Value > 1.5 Then
            Sheets("d_P").Cells(jj, ii).Value = Sheets("d_P").Cells(jj, ii + 2).Value / Sheets("d_P").Cells(jj + 1, ii + 2).Value
            Sheets("d_P").Cells(jj, ii).Interior.ColorIndex = 7
        Else
            Sheets("d_P").Cells(jj, ii).Value = Sheets("d_P").Cells(jj, ii + 2).Value / Sheets("d_P").Cells(jj + 1, ii + 2).Value
        End If
    Next
Next
End Sub

