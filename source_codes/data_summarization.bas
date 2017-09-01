Attribute VB_Name = "data_summarization"
Option Explicit

'更新时间: 2014/04/02 14:53
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:
'1.Label1后一行添加激活模型调试汇总表
'2.平动系数间隔符用“~”


'更新时间: 2014/04/02 14:53
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:
'1.更改general表名
'2.增加读取wmass
'3.周期之间间隔符用“~”


'更新时间: 2013/7/30 20:26
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:
'1.添加时间列，如general中没有时间数据则读取当前时间日期

'更新时间: 2013/7/23 23:26
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:
'1.更新读取汇总表方式为两种：1，读取文件夹内OUT文件然后读入；2，直接读取已有OUTReader文件内数据。均可独立运行，不需在查看模式中进行操作


'更新时间: 2013/7/20 19:06
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:
'1.读取任意工作簿中geneeral表数据到指定工作簿中模型调试汇总表，读取数据时，数据表会被打开

'更新时间: 2013/7/20 19:06
'///////////////////////////////////////////////////////////////////////////////////////////////////////
'更新内容:

Public Sub Data_Summarization(WB_Orig_Full As String, WB_New_Full As String, general As String)


'计算运行时间
Dim sngStart As Single
sngStart = Timer

'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                       工作表"模型调试汇总"设定                       ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************


'==============================================================================打开模型调试workbook

'定义工作簿名称
Dim WB_New As String '汇总表
Dim WB_Orig As String '数据表

Application.DisplayAlerts = False

'从完整路径中提取原始数据工作簿名称
Dim aFile_1 As Variant
aFile_1 = Split(WB_Orig_Full, "\")
WB_Orig = aFile_1(UBound(aFile_1))

'从完整路径中提取汇总数据工作簿名称
Dim aFile_2 As Variant
aFile_2 = Split(WB_New_Full, "\")
WB_New = aFile_2(UBound(aFile_2))

'如果原始数据工作簿未打开，则打开工作簿
Debug.Print WB_Orig
If Not WorkbookOpen(WB_Orig) Then
    Workbooks.Open (WB_Orig_Full)
Else
    MsgBox "模型数据工作簿已打开"
End If

'如果汇总数据工作簿未打开，则打开工作簿
If Not WorkbookOpen(WB_New) Then
    Workbooks.Open (WB_New_Full)
Else
    MsgBox "汇总数据工作簿已打开"
End If

Workbooks(WB_New).Activate

Dim Wb As Worksheet

'搜寻已有的工作表的名称
For Each Wb In Worksheets
    '汇总工作表存在则跳出工作表格式设定
    If Wb.name = "模型调试汇总" Then
        GoTo Label1
    End If
Next Wb
    
'==============================================================================添加表格"模型调试汇总"的标题

Debug.Print "开始设定表格""模型调试汇总""的格式"
Debug.Print "……"

'新建工作表
Call Addsh("模型调试汇总")

With Sheets("模型调试汇总")

'清除工作表所有内容
.Cells.Clear

'调整单元格宽/高
.range("A:A").ColumnWidth = 4
.range("B:F,L:L").ColumnWidth = 8
.range("M:N").ColumnWidth = 9
.range("G:I").ColumnWidth = 15
.range("J:K,O:Q").ColumnWidth = 12
.Rows("1:1").RowHeight = 35

End With

'设置纸张大小，A3横向
Sheets("模型调试汇总").PageSetup.PaperSize = xlPaperA3
Sheets("模型调试汇总").PageSetup.Orientation = xlLandscape

'设为分页视图
Sheets("模型调试汇总").Activate
ActiveWindow.View = xlPageBreakPreview
ActiveWindow.Zoom = 90

'加表格线
Call AddFormLine("模型调试汇总", "A1:Q100")

'加背景色
Call AddShadow("模型调试汇总", "A2:I2", 16777164)
Call AddShadow("模型调试汇总", "J2:Q2", 6750105)

Debug.Print "设定表格""模型调试汇总""的格式完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"


Debug.Print "开始添加表格""模型调试汇总""的标题"
Debug.Print "……"

'------------------------------------------------------工作表""模型调试汇总""内的标题格式
With Sheets("模型调试汇总")
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '设置小数点后位数
    '.Cells.NumberFormatLocal = "0.00"
    '设置局部单元格特殊格式
    .range("B:B").NumberFormatLocal = "yyyy/m/d"
    'Cells(14, 7).NumberFormatLocal = "# ???/???"
    'Cells(15, 7).NumberFormatLocal = "G/通用格式"
    'Cells(17, 7).NumberFormatLocal = "G/通用格式"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = True
    
    '-------------------------------------------------表头区
    '合并单元格
    .range("A1:Q1").MergeCells = True
    .Cells(1, 1).HorizontalAlignment = xlCenter
    .Cells(1, 1).VerticalAlignment = xlCenter
    
    '表头
    .Cells(1, 1).Font.name = "黑体"
    .Cells(1, 1).Font.Size = "20"
    .Cells(1, 1) = "模型调试汇总记录表"
    
    '-------------------------------------------------项目信息区
    '项目信息
    .Cells(2, 1) = "No."
    .Cells(2, 2) = "时间"
    .Cells(2, 3) = "调整阶段"
    .Cells(2, 4) = "模型"
    .Cells(2, 5) = "原始模型"
    .Cells(2, 6) = "文件夹"
    .Cells(2, 7) = "目标"
    .Cells(2, 8) = "操作"
    .Cells(2, 9) = "结果"
    .Cells(2, 10) = "T1T2T3(s)"
    .Cells(2, 11) = "平动系数"
    .Cells(2, 12) = "Tt/T1"
    .Cells(2, 13) = "质量系数X"
    .Cells(2, 14) = "质量系数Y"
    .Cells(2, 15) = "层间位移角"
    .Cells(2, 16) = "最大位移比"
    .Cells(2, 17) = "层间位移比"
    
End With

'保存
ActiveWorkbook.Save

Debug.Print "添加表格""模型调试汇总""的标题完毕"
Debug.Print "运行时间: " & Timer - sngStart
Debug.Print "……"
    


'******************************************************************************
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                        读取工作表general内容                         ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************
Label1:
Workbooks(WB_New).Sheets("模型调试汇总").Activate

Dim N_Model As Integer

'确定已有数据行数
N_Model = ActiveSheet.range("A65535").End(xlUp).Row

'序号
Sheets("模型调试汇总").Cells(N_Model + 1, 1) = N_Model - 1

'时间  如果general中没有时间日期，读取当前日期
If IsEmpty(Workbooks(WB_Orig).Sheets(general).Cells(4, 7)) Then
    Sheets("模型调试汇总").Cells(N_Model + 1, 2) = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
Else
    Sheets("模型调试汇总").Cells(N_Model + 1, 2) = Workbooks(WB_Orig).Sheets(general).Cells(4, 7)
End If

'调整阶段
Sheets("模型调试汇总").Cells(N_Model + 1, 3) = Information_Input.TextBox_Stage.Text

'模型
Sheets("模型调试汇总").Cells(N_Model + 1, 4) = Information_Input.TextBox_Model.Text

'原始模型
Sheets("模型调试汇总").Cells(N_Model + 1, 5) = Information_Input.TextBox_Orig_Model.Text

'文件夹
Sheets("模型调试汇总").Cells(N_Model + 1, 6) = Information_Input.TextBox_Folder.Text

'目标
Sheets("模型调试汇总").Cells(N_Model + 1, 7) = Information_Input.TextBox_Target.Text

'序号
Sheets("模型调试汇总").Cells(N_Model + 1, 8) = Information_Input.TextBox_Operate.Text
'序号
Sheets("模型调试汇总").Cells(N_Model + 1, 9) = Information_Input.TextBox_Result.Text

'周期 T1T2T3
Sheets("模型调试汇总").Cells(N_Model + 1, 10) = Round(Workbooks(WB_Orig).Sheets(general).Cells(28, 4), 2) & "~" & Round(Workbooks(WB_Orig).Sheets(general).Cells(29, 4), 2) & "~" & Round(Workbooks(WB_Orig).Sheets(general).Cells(30, 4), 2)

'平动系数
Sheets("模型调试汇总").Cells(N_Model + 1, 11) = Round(1 - Workbooks(WB_Orig).Sheets(general).Cells(28, 7), 2) & "~" & Round(1 - Workbooks(WB_Orig).Sheets(general).Cells(29, 7), 2) & "~" & Round(1 - Workbooks(WB_Orig).Sheets(general).Cells(30, 7), 2)

'Tt / T1
Sheets("模型调试汇总").Cells(N_Model + 1, 12) = Round(Workbooks(WB_Orig).Sheets(general).Cells(38, 4), 2)

'质量系数X
Sheets("模型调试汇总").Cells(N_Model + 1, 13) = Workbooks(WB_Orig).Sheets(general).Cells(39, 5) & "%"

'质量系数Y
Sheets("模型调试汇总").Cells(N_Model + 1, 14) = Workbooks(WB_Orig).Sheets(general).Cells(39, 7) & "%"

'层间位移角
Sheets("模型调试汇总").Cells(N_Model + 1, 15) = Workbooks(WB_Orig).Sheets(general).Cells(14, 4) & "(" & Workbooks(WB_Orig).Sheets(general).Cells(15, 5) & ")"

'最大位移比
Sheets("模型调试汇总").Cells(N_Model + 1, 16) = Workbooks(WB_Orig).Sheets(general).Cells(16, 4) & "(" & Workbooks(WB_Orig).Sheets(general).Cells(17, 5) & ")"

'层间位移比
Sheets("模型调试汇总").Cells(N_Model + 1, 17) = Workbooks(WB_Orig).Sheets(general).Cells(18, 4) & "(" & Workbooks(WB_Orig).Sheets(general).Cells(19, 5) & ")"

'保存
ThisWorkbook.Save

Information_Input.Hide
OUTReader_Main.Show 0

'单元格宽度自动调整
'Sheets("模型调试汇总").Cells.EntireRow.AutoFit

MsgBox "耗费时间: " & Timer - sngStart

End Sub

Function WorkbookOpen(WorkBookName As String) As Boolean
    '如果该工作簿已打开则返回真
    WorkbookOpen = False
    On Error GoTo WorkBookNotOpen
    If Len(Application.Workbooks(WorkBookName).name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkBookNotOpen:
End Function

 
Public Sub Test_XX(XXC As Integer)

XXC = 1
Dim WB_New_Full As String
Dim WB_Orig_Full As String
'WB_Orig_Full = OUTReader_Main.TextBox_Path_2.Text
WB_Orig_Full = ThisWorkbook.Path & "\" & ThisWorkbook.name
WB_New_Full = OUTReader_Main.TextBox_Path_3.Text
'Path = Workbooks(WB_Orig).Worksheets("debug").Cells(2, 2)
'Debug.Print Path
If OUTReader_Main.CheckBox_PKPM_2 Then
    Call Data_Summarization(WB_Orig_Full, WB_New_Full, "g_P")
End If

If OUTReader_Main.CheckBox_YJK_2 Then
    Call Data_Summarization(WB_Orig_Full, WB_New_Full, "g_Y")
End If

If OUTReader_Main.CheckBox_MBuilding_2 Then
    Call Data_Summarization(WB_Orig_Full, WB_New_Full, "g_M")
End If

End Sub

