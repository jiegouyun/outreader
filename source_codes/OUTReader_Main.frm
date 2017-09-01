VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OUTReader_Main 
   Caption         =   "Main"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7200
   OleObjectBlob   =   "OUTReader_Main.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "OUTReader_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/4/29
'1.修正剪重比限值的bug


'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/4//19
'1.添加ETABS高亮



'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/7/05
'1.修改提取支撑内力时的错误


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/19
'1.对构件内力-选择路径，添加代码
'2.修改“模型对比”中“三选二”
'3.修改发布日期


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'1.将模型路径文件放在D盘下。
'2.添加ETABS与其他三个软件的对比。


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/18
'1.添加路径写出及读入代码


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/2
'1.调试模式增加未选计算软件的提示


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/28
'1.添加了ETABS V9的选项

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/24
'1.增加路径框默认值

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/18
'1.刚度比修正按钮增加YJK控制


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/8
'1.增加ETABS数据提取页面及代码，在最下部


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/3
'1.更新读取文件路径语句，默认打开当前文件所在文件夹


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/9
'1.新增构件内力板面，用于提取构件的标准内力、组合工况并校核等


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/7
'1.点击打开文件夹路径时打开当前文件夹

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/5
'1.修改“模型对比”下的信息选项，使限值能覆盖全部范围，代码略繁琐，如果用数组应该会好些，以后可改进；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/1/4
'1.对应剪力墙受剪信息提取模块修改按钮布置
'2.增加读取墙编号总数的按钮
'3.删除手工输入分布筋部分
'4.关于板面中的删除所有表格，保留隐藏的墙配筋表格wallrebar

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/12/11
'1.提取剪力墙受剪信息部分分两个步骤进行，可以修改墙编号，再进行数据读取

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/27
'1.更新画图程序。（我在公司可以正常使用，在家里使用时点击生成反应谱曲线会弹出一个路径选择框，取消后不影响后面运行）

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/25
'1.添加关于栏,调整数据对比代码;
'2.添加删除所有工作表按钮（藏起来了，看需不需要提供使用）和快捷键alt+m；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/16
'1.将表格生成和图表生成分开
'2.表名简化,如ColumnInfo改为CI
'3.配筋对比增加PKPM与YJK选项，其中PKPM默认选中，YJK暂时锁定

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/12
'1.简化表名，如general_PKPM:d_P，distribution_YJK:d_Y等。
'2.删除界面上的说明

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/8/4

'更新内容：
'1.增加路径写入代码；


'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/30 21:30

'更新内容：
'1.调试模式中工作簿路径初始值改为工作簿所在文件夹路径，与OUT默认情况对应


'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/29 19:50

'更新内容：
'1.添加MBuilding模块
'2.添加配筋对比模块
'3.修改了调试模式界面某些文字
'4.修正全楼层下首层轴压比错误

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/23 13:50

'更新内容：
'1.修改快捷键设置，现在快捷键已可以使用
'2.添加模型数据表选取路径窗口，可以直接调用其它工作簿中general表中数据
'3.添加路径框初始默认值，查看模式中数据路径为工作簿所在路径，调试模式模型数据表和汇总数据表默认为程序所在工作簿。


'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/20 21:50

'更新内容：
'1.加入模型调试汇总内容，但是不知道哪里出了问题，从窗体中直接调用Data_Summarization模块总是出现错误，你们看一下是什么原因，目前是通过一个中间模块转换了一次。

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/16 18:50

'更新内容：
'1.更改写入新general的轴压比代码

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/7/3 20:21

'更新内容：
'1.解锁YJK选项
'2.绘图区分YJK和PKPM

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/6/27 22：23

'更新内容：
'1.将WPJ改为复选框

'////////////////////////////////////////////////////////////////////////////
'更新时间:2013/6/27

'更新内容：
'1.添加DYNA及DYNA绘图；
'2.解锁WMASS；
'3.添加


'////////////////////////////////////////////////////////////////////////////


'更新时间:2013/6/06

'更新内容：
'1.WPJ功能分拆；
'2.调整窗体出现的位置顺序；
'3.调整临时位移角数据的位置，设为白色并隐藏；
'4.添加说明；

'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/05/

'更新内容：
'1.为窗体按钮添加快捷键
'2.增加绘图功能


'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/04/

'更新内容：
'1.已现有代码更改了工作表名称，进行测试，后面可以将工作表名称也设置为变量，便于重新命名；
'2.将Main模块中的一些代码移植到了窗体中，如select、zoom、"将首层墙柱最大轴压比及其构件编号写入general"；
'3.添加了退出按钮，大家看需不需要；
'4.对复选框的默认状态进行了设置，原则是默认全选；Wmass文件由于要输出层数等总体信息，默认选中并将其锁定；YJK，Dyna文件还在完善，默认未选，并将其锁定，以后完善后解锁；
'5.在模块中增加一模块，给窗体设置打开的快捷键F2，关闭窗体后可用F2调出窗体；相应在ThisWorkbook里添加了相应代码；
'6.增加选择文件夹路径按钮；
'////////////////////////////////////////////////////////////////////////////

'更新时间:2013/6/01/ 17:57


'////////////////////////////////////////////////////////////////////////////



'==========================================================================================配筋对比图表窗口
'配筋对比图表生成按钮
Private Sub CompareS_Figure_Click()
'计算运行时间
Dim start As Single
start = Timer

OUTReader_Main.Hide


If CheckBox6_PKPM Then
    '调用子程序
    Call OUTReader_PKPM_CompareS_figure
End If
OUTReader_Main.Show 0

MsgBox "耗费时间: " & Timer - start

End Sub



'==========================================================================================配筋对比表格窗口
'配筋对比表格生成按钮
Private Sub CompareS_tabale_Click()
'计算运行时间
Dim start As Single
start = Timer

OUTReader_Main.Hide

'定义主要辅助变量

Dim path1 As String, path2 As String
Dim startf As Integer, endf As Integer

'读取目录路径
path1 = OUTReader_Main.TextBox_path_xz.Text
path2 = OUTReader_Main.TextBox_path_zz.Text
'读取层号
startf = OUTReader_Main.TextBox_startf.Text
endf = OUTReader_Main.TextBox_endf.Text

'调用子程序
If CheckBox6_PKPM Then

    Call OUTReader_PKPM_CompareS_table(path1, path2, startf, endf)
    
End If

OUTReader_Main.Show 0

MsgBox "耗费时间: " & Timer - start

End Sub

Private Sub AllMem_CommandButton5_Click()

Mem_S.Value = 1
If soft_YJK Then
    Mem_E.Value = allmem_Y(Dic_TextBox.Text, F_Num_TextBox.Value)
    Else: If soft_PKPM Then Mem_E.Value = allmem_P(Dic_TextBox.Text, F_Num_TextBox.Value)
End If

End Sub





'=============================================================================单片墙剪力校核_生成表格
Private Sub FormCreate_Click()
'模型表格按钮
'生成表格格式
'计算运行时间
Dim start As Single
start = Timer


Dim path1 As String
path1 = Dic_TextBox.Text

'生成PKPM的墙剪力校核表格
If soft_PKPM Then

    If Refer_F Then
        Debug.Print path1
        Call member_info_f(path1, M_Num_TextBox.Value, FLoor_S.Value, Floor_E.Value, "P", "M" & M_Num_TextBox.Value)
        MsgBox "请修改墙编号后，运行【数据读取】"
    End If
    
    If Refer_M Then
        Debug.Print path1
        Call member_info_m(path1, F_Num_TextBox.Value, Mem_S.Value, Mem_E.Value, "P", "F" & F_Num_TextBox.Value)
    End If
    
End If

'生成YJK的墙剪力校核表格
If soft_YJK Then

    If Refer_F Then
        Debug.Print path1
        Call member_info_f(path1, M_Num_TextBox.Value, FLoor_S.Value, Floor_E.Value, "Y", "M" & M_Num_TextBox.Value)
        MsgBox "请修改墙编号后，运行【数据读取】"
    End If
    
    If Refer_M Then
        Debug.Print path1
        Call member_info_m(path1, F_Num_TextBox.Value, Mem_S.Value, Mem_E.Value, "Y", "F" & F_Num_TextBox.Value)
    End If
    
End If


'Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"
'Sheets("d_M").Columns("B:C").NumberFormatLocal = "0.00"

OUTReader_Main.Show 0

'ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start

End Sub

'=============================================================================构件内力文件夹路径读取
Private Sub Get_Dic_MD_Click()

Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then MD_path.Text = fd.SelectedItems(1)

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame23_Click()

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Label54_Click()

End Sub

Private Sub Label51_Click()

End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label9_Click()

End Sub

'=============================================================================组合工况提取_读取数据
Private Sub LoadComb_Click()

'计算运行时间
Dim start As Single
start = Timer

If MD_YJK Then

    Call LOADCOMB_WC_Y(MD_path.Text)
    
End If


If MD_PKPM Then

    Call LOADCOMB_WC_P(MD_path.Text)
    
End If


MsgBox "耗费时间: " & Timer - start

End Sub


Private Sub MD_BEAM_CheckBox_Click()

End Sub

Private Sub MD_WC_CheckBox_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

'=============================================================================构件内力校核

Private Sub SigleMem_check_Click()

'计算运行时间
Dim start As Single
start = Timer

Dim Path As String
Path = MD_path.Text

Dim shname As String

If MD_YJK Then
     
     If MD_WC_CheckBox Then
        shname = "Y_WCC_F" & MD_FLO2.Value
        Call SingleWallData_Y(Path, shname, MD_FLO2.Value, MD_NUM.Value)
    End If
    
     If MD_C_CheckBox Then
        shname = "Y_CC_F" & MD_FLO2.Value
        Call SingleColData_Y(Path, shname, MD_FLO2.Value, MD_NUM.Value, "C")
    End If
    
     If MD_GC_CheckBox Then
        shname = "Y_GC_F" & MD_FLO2.Value
        Call SingleColData_Y(Path, shname, MD_FLO2.Value, MD_NUM.Value, "G")
    End If
    
End If


If MD_PKPM Then
     
     If MD_WC_CheckBox Then
        shname = "P_WCC_F" & MD_FLO2.Value
    End If
    
    Call SingleWallData_P(Path, shname, MD_FLO2.Value, MD_NUM.Value)
    
End If


'ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start

End Sub


'=============================================================================构件内力提取_读取数据

Private Sub MD_ALL_Click()

'计算运行时间
Dim start As Single
start = Timer

Dim Path As String
Path = MD_path.Text

Dim shname As String


If MD_YJK Then
     
     If MD_WC_CheckBox Then
        shname = "Y_WCD_F" & MD_FLO.Value
        Call WallData_Y(Path, shname, MD_FLO.Value)
    End If
    
    If MD_C_CheckBox Then
        shname = "Y_CD_F" & MD_FLO.Value
        Call ColData_Y(Path, shname, MD_FLO.Value, "C")
    End If
    
    If MD_GC_CheckBox Then
        shname = "Y_GD_F" & MD_FLO.Value
        Call ColData_Y(Path, shname, MD_FLO.Value, "G")
    End If
    
    If MD_BEAM_CheckBox Then
        shname = "Y_BEAM_F" & MD_FLO.Value
        Call BeamData_Y(Path, shname, MD_FLO.Value, "G")
    End If
    
End If


If MD_PKPM Then
     
     If MD_WC_CheckBox Then
        shname = "P_WCD_F" & MD_FLO.Value
        Call WallData_P(Path, shname, MD_FLO.Value)
    End If
    
'    If MD_C_CheckBox Then
'        shname = "P_CD_F" & MD_FLO.Value
'        Call ColData_P(Path, shname, MD_FLO.Value, "C")
'    End If
'
'    If MD_GC_CheckBox Then
'        shname = "P_GD_F" & MD_FLO.Value
'        Call ColData_P(Path, shname, MD_FLO.Value, "G")
'
'    End If
    
End If

'ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start

End Sub



'=============================================================================单片墙剪力校核_读取数据

Private Sub ReadData_Click()
'模型表格按钮
'数据读取

'计算运行时间
Dim start As Single
start = Timer


Dim path1 As String
path1 = Dic_TextBox.Text


Dim i As Integer

If soft_PKPM Then

    If Refer_F Then
        For i = CStr(FLoor_S.Value) To CStr(Floor_E.Value)
            Call PKPM_Wall_Info_F(path1, i, "P", "M" & M_Num_TextBox.Value)
        Next
    End If
    
    If Refer_M Then
        For i = CStr(Mem_S.Value) To CStr(Mem_E.Value)
            Call PKPM_Wall_Info_M(path1, i, "P", "F" & F_Num_TextBox.Value)
        Next
    End If
    
End If

If soft_YJK Then

    If Refer_F Then
        For i = CStr(FLoor_S.Value) To CStr(Floor_E.Value)
            Call YJK_Wall_Info_F(path1, i, "Y", "M" & M_Num_TextBox.Value)
        Next
    End If
    
    If Refer_M Then
        For i = CStr(Mem_S.Value) To CStr(Mem_E.Value)
            Call YJK_Wall_Info_M(path1, i, "Y", "F" & F_Num_TextBox.Value)
        Next
    End If
    
End If


'Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"
'Sheets("d_M").Columns("B:C").NumberFormatLocal = "0.00"

OUTReader_Main.Show 0

'ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start
End Sub

'=============================================================================单片墙剪力校核_生成表格

'模型表格按钮
Private Sub FigureData_Click()

'计算运行时间
Dim start As Single
start = Timer

If soft_PKPM Then

    If Refer_F Then
        Call member_wall_f("P", "M" & M_Num_TextBox.Value)
    End If
    
    If Refer_M Then
        Call member_wall_m("P", "F" & F_Num_TextBox.Value)
    End If
    
End If

If soft_YJK Then

    If Refer_F Then
        Call member_wall_f("Y", "M" & M_Num_TextBox.Value)
    End If
    
    If Refer_M Then
        Call member_wall_m("Y", "F" & F_Num_TextBox.Value)
    End If
    
End If
    

'Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"
'Sheets("d_M").Columns("B:C").NumberFormatLocal = "0.00"

OUTReader_Main.Show 0

'ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start

End Sub



Private Sub TextBox_Path_2_Change()

End Sub

'为窗体按钮添加快捷键 奇怪这么不成功，就先在属性里加了.已可以使用
Private Sub UserForm_Initialize()
    'page_1.Accelerator = "1" '选择路径：ALT+1
    'Page_2.Accelerator = "2" '选择路径：ALT+2
    OUTReader_Main.Get_Dic.Accelerator = "a" '选择路径：ALT+A
    OUTReader_Main.Get_dic_2.Accelerator = "o" '选择模型数据工作簿：ALT+O
    OUTReader_Main.Get_dic_3.Accelerator = "n" '选择汇总工作簿：ALT+N
    OUTReader_Main.Get_Data.Accelerator = "d" '数据表格：ALT+D
    OUTReader_Main.Get_Figure.Accelerator = "c" '生成图表：ALT+C
    OUTReader_Main.Get_Figure_Dyna.Accelerator = "t" '生成时程图表：ALT+T
    OUTReader_Main.QuitButton.Accelerator = "q" '退出程序：ALT+Q
    OUTReader_Main.Data_Summarize.Accelerator = "s" '汇总数据：ALT+S
    OUTReader_Main.dsheets.Accelerator = "m" '删除所有工作表：ALT+m
    
    OUTReader_Main.TextBox_Path.Text = ThisWorkbook.Path '工作路径初始默认值
    OUTReader_Main.TextBox_Path_2.Text = ThisWorkbook.Path  '模型数据工作路径初始默认值
    OUTReader_Main.TextBox_Path_3.Text = ThisWorkbook.Path & "\" & ThisWorkbook.name '模型汇总工作簿初始默认值为本xlsm
    OUTReader_Main.TextBox_path_xz.Text = ThisWorkbook.Path '小震工作路径初始默认值
    OUTReader_Main.TextBox_path_zz.Text = ThisWorkbook.Path '中震工作路径初始默认值
    OUTReader_Main.Dic_TextBox.Text = ThisWorkbook.Path '墙受剪验算工作路径初始默认值
    OUTReader_Main.MD_path.Text = ThisWorkbook.Path '构件内力工作路径初始默认值
    OUTReader_Main.TextBox_Path_ETABS.Text = ThisWorkbook.Path 'ETABS工作路径初始默认值
    
  
End Sub

'选择目录按钮
Private Sub Get_Dic_31_Click()
Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then TextBox_path_xz.Text = fd.SelectedItems(1)
End Sub

'选择目录按钮
Private Sub Get_Dic_32_Click()
Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then TextBox_path_zz.Text = fd.SelectedItems(1)
End Sub

'选择目录按钮
Private Sub Get_Dic_6_Click()

Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)

If fd.Show = -1 Then
Dic_TextBox.Text = fd.SelectedItems(1)
End If

If Not Sheets("d_P").Cells(3, 1) = "" Then
    Num_all = Sheets("d_P").range("a65536").End(xlUp)
    Floor_E.Value = Num_all
End If

End Sub


'选择目录按钮
Private Sub CommandButton9_Click()

Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then
MD_path.Text = fd.SelectedItems(1)
End If


End Sub




'=============================================================================调试窗体

'生成调试模型汇总信息
Private Sub Data_Summarize_Click()
If OUTReader_Main.CheckBox_PKPM_2.Value = False And OUTReader_Main.CheckBox_YJK_2.Value = False And OUTReader_Main.CheckBox_MBuilding_2.Value = False Then
    MsgBox ("请选择计算软件！")
    Exit Sub
End If

OUTReader_Main.Hide
Information_Input.Show 0

End Sub

'调试模式数据表选择
Private Sub Get_dic_2_Click()
Dim fd As FileDialog
'从Excel中读取数据
If OptionButton_Excel Then
    Application.FileDialog(msoFileDialogFilePicker).InitialFileName = ThisWorkbook.Path & "\"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xls*"
    If fd.Show = -1 Then OUTReader_Main.TextBox_Path_2 = fd.SelectedItems(1)
'    Dim WBToWrite As String
'    WBToWrite = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm")
'    If WBToWrite <> "False" Then
'        OUTReader_Main.TextBox_Path_2 = WBToWrite
'    End If
End If

'从OUT数据中读取数据
If OptionButton_OUT Then
    Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then TextBox_Path_2.Text = fd.SelectedItems(1)
End If
End Sub

'调试模式汇总表选择
Private Sub Get_dic_3_Click()
    Dim fd As FileDialog
    Application.FileDialog(msoFileDialogFilePicker).InitialFileName = ThisWorkbook.Path & "\"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xls*"
    If fd.Show = -1 Then OUTReader_Main.TextBox_Path_3 = fd.SelectedItems(1)
    'Dim WBToWrite As String
    'WBToWrite = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    'If WBToWrite <> "False" Then
    '    OUTReader_Main.TextBox_Path_3 = WBToWrite
'End If
End Sub


'=============================================================================查看窗体

'选择目录按钮
Private Sub Get_Dic_Click()
Dim fd As FileDialog
Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then TextBox_Path.Text = fd.SelectedItems(1)
End Sub

'生成表格按钮
Private Sub Get_Data_Click()
'On Error Resume Next
'计算运行时间
Dim start As Single
start = Timer


OUTReader_Main.Hide

'==========================================================================================定义主要辅助变量

Dim Path As String
Dim general, Distribution, Program As String


'--------------------------------------------------------------删除原有数据
'Dim i_s As Integer
'For i_s = Sheets.Count To 1 Step -1
    'Sheets(i_s).Cells.Clear
'Next


'--------------------------------------------------------------获取数据存储路径
Path = OUTReader_Main.TextBox_Path.Text


'==========================================================================================提取PKPM数据
If CheckBox_PKPM Then

 Open "d:\dic.ini" For Output As #1
        Print #1, Path
    Close #1

    '----------------------------------------------------------生成表格表头
    'Program = "PKPM"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_P")
    Call Addsh("d_P")
    Call Addsh("CR_P")
    Call Addsh("WR_P")
    Call AddHeadline("g_P", "d_P", "CR_P", "WR_P")
    Sheets("g_P").Cells(3, 4) = Path


    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)
    If CheckBox_WMASS Then
        'OptionButton_WPJF.Locked = True
        'OptionButton_WPJA.Locked = True
        Sheets("d_P").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_PKPM_WMASS(Path)
    End If

    '----------------------------------------------------------写入wzq.out文件数据
    If CheckBox_WZQ Then
        Sheets("d_P").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_PKPM_WZQ(Path)
    End If

    '----------------------------------------------------------写入wdisp.out文件数据
    If CheckBox_WDISP Then
        Sheets("d_P").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_PKPM_WDISP(Path)
    End If

    '----------------------------------------------------------写入wpj.out文件数据
    If CheckBox_WPJF Then
        Sheets("CR_P").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_P").Select
        ActiveWindow.Zoom = 70
        Call OUTReader_PKPM_WPJ_UC(Path, Num_Base + 1)
    '------------------------------------将首层墙柱最大轴压比及其构件编号写入general
        Sheets("g_P").Cells(8, 5) = Sheets("d_P").Cells(Num_Base + 3, 56)
        Sheets("g_P").Cells(8, 7) = Round(Sheets("d_P").Cells(Num_Base + 3, 57), 1)
        Sheets("g_P").Cells(9, 5) = Sheets("d_P").Cells(Num_Base + 3, 58)
        Sheets("g_P").Cells(9, 7) = Round(Sheets("d_P").Cells(Num_Base + 3, 59), 1)
    End If
    If CheckBox_WPJA Then
        Sheets("CR_P").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_P").Select
        ActiveWindow.Zoom = 70
        Dim i As Integer
        For i = 1 To Num_all
            Call OUTReader_PKPM_WPJ_UC(Path, i)
        Next
    '------------------------------------将全楼层墙柱最大轴压比及其构件编号写入general
        Sheets("g_P").Cells(8, 5) = Sheets("d_P").Cells(Num_Base + 3, 56)
        Sheets("g_P").Cells(8, 7) = Round(Sheets("d_P").Cells(Num_Base + 3, 57), 1)
        Sheets("g_P").Cells(9, 5) = Sheets("d_P").Cells(Num_Base + 3, 58)
        Sheets("g_P").Cells(9, 7) = Round(Sheets("d_P").Cells(Num_Base + 3, 59), 1)
    End If



    '----------------------------------------------------------写入wv02q.out文件数据
    If CheckBox_WV02Q Then
        Sheets("d_P").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_PKPM_WV02Q(Path)
    End If

    '----------------------------------------------------------写入wdyna.out文件数据
    If CheckBox_WDYNA Then
        Call Addsh("e_P")
        Sheets("e_P").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_PKPM_WDYNA(Path)
    End If
    
    Call gaoliang("P")
    
    '设置小数点后位数
    Sheets("d_P").Cells.NumberFormatLocal = "G/通用格式"
    Sheets("d_P").Columns("B:C").NumberFormatLocal = "0.00"
    Sheets("d_P").Columns("AT:AU").NumberFormatLocal = "0.00"
    Sheets("d_P").Columns("AH:AS").NumberFormatLocal = "0.00"
    Sheets("d_P").Columns("BB:BB").NumberFormatLocal = "0.0"
    Sheets("d_P").Columns("BC:BC").NumberFormatLocal = "0.00"
    
    Sheets("g_P").Select

' Open "d:\dic.ini" For Output As #1
'        Print #1, Path
'    Close #1

End If

'==========================================================================================提取YJK数据
If CheckBox_YJK Then

 Open "d:\dic.ini" For Output As #1
        Print #1, Path
    Close #1

    '----------------------------------------------------------生成表格表头
    'Program = "YJK"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_Y")
    Call Addsh("d_Y")
    Call Addsh("CR_Y")
    Call Addsh("WR_Y")
    Call AddHeadline("g_Y", "d_Y", "CR_Y", "WR_Y")
    Sheets("g_Y").Cells(3, 4) = Path
    
        
    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)
    If CheckBox_WMASS Then
        'OptionButton_WPJF.Locked = True
        'OptionButton_WPJA.Locked = True
        Sheets("d_Y").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_YJK_WMASS(Path)
    End If
        
    '----------------------------------------------------------写入wzq.out文件数据
    If CheckBox_WZQ Then
        Sheets("d_Y").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_YJK_WZQ(Path)
    End If
        
    '----------------------------------------------------------写入wdisp.out文件数据
    If CheckBox_WDISP Then
        Sheets("d_Y").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_YJK_WDISP(Path)
    End If
    
    '----------------------------------------------------------写入wpj.out文件数据
    If CheckBox_WPJF Then
        Sheets("CR_Y").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_Y").Select
        ActiveWindow.Zoom = 70
        Call OUTReader_YJK_WPJ_UC(Path, Num_Base + 1)
    '------------------------------------将首层墙柱最大轴压比及其构件编号写入general
        Sheets("g_Y").Cells(8, 5) = Sheets("d_Y").Cells(Num_Base + 3, 56)
        Sheets("g_Y").Cells(8, 7) = Round(Sheets("d_Y").Cells(Num_Base + 3, 57))
        Sheets("g_Y").Cells(9, 5) = Sheets("d_Y").Cells(Num_Base + 3, 58)
        Sheets("g_Y").Cells(9, 7) = Round(Sheets("d_Y").Cells(Num_Base + 3, 59))
    End If
    If CheckBox_WPJA Then
        Sheets("CR_Y").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_Y").Select
        ActiveWindow.Zoom = 70
        For i = 1 To Num_all
            Call OUTReader_YJK_WPJ_UC(Path, i)
        Next
    '------------------------------------将全楼层墙柱最大轴压比及其构件编号写入general
        Sheets("g_Y").Cells(8, 5) = Sheets("d_Y").Cells(Num_Base + 3, 56)
        Sheets("g_Y").Cells(8, 7) = Round(Sheets("d_Y").Cells(Num_Base + 3, 57))
        Sheets("g_Y").Cells(9, 5) = Sheets("d_Y").Cells(Num_Base + 3, 58)
        Sheets("g_Y").Cells(9, 7) = Round(Sheets("d_Y").Cells(Num_Base + 3, 59))
    End If
 
    '----------------------------------------------------------写入wv02q.out文件数据
    If CheckBox_WV02Q Then
        Sheets("d_Y").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_YJK_WV02Q(Path)
    End If
    
    '----------------------------------------------------------写入wdyna.out文件数据
    If CheckBox_WDYNA Then
        Call Addsh("e_YJK")
        Sheets("e_YJK").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_YJK_WDYNA(Path)
    End If
    
     Call gaoliang("Y")

'    Open "d:\dic.ini" For Output As #1
'        Print #1, Path
'    Close #1
    
    '设置小数点后位数
    Sheets("d_Y").Cells.NumberFormatLocal = "G/通用格式"
    Sheets("d_Y").Columns("B:C").NumberFormatLocal = "0.00"
    Sheets("d_Y").Columns("AT:AU").NumberFormatLocal = "0.00"
    Sheets("d_Y").Columns("AH:AS").NumberFormatLocal = "0.00"
    
    Sheets("g_Y").Select
        
End If


'==========================================================================================提取MBuilding数据

If CheckBox_MBuilding Then

 Open "d:\dic.ini" For Output As #1
        Print #1, Path
    Close #1

    '----------------------------------------------------------生成表格表头
    'Program = "MBuilding"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_M")
    Call Addsh("d_M")
    Call Addsh("CR_M")
    Call Addsh("WR_M")
    Call AddHeadline("g_M", "d_M", "CR_M", "WR_M")
    Sheets("g_M").Cells(3, 4) = Path
    
        
    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)
    If CheckBox_WMASS Then
        'OptionButton_WPJF.Locked = True
        'OptionButton_WPJA.Locked = True
        Sheets("d_M").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_MBuilding_总信息(Path)
        Call OUTReader_MBuilding_侧向刚度(Path)
        Call OUTReader_MBuilding_抗剪承载力(Path)
    End If
        
    '----------------------------------------------------------写入wzq.out文件数据
    If CheckBox_WZQ Then
        Sheets("d_M").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_MBuilding_周期振型(Path)
    End If
        
    '----------------------------------------------------------写入wdisp.out文件数据
    If CheckBox_WDISP Then
        Sheets("d_M").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_MBuilding_结构位移(Path)
    End If
    
    '----------------------------------------------------------写入wpj.out文件数据
    If CheckBox_WPJF Then
        Sheets("CR_M").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_M").Select
        ActiveWindow.Zoom = 70
        Call OUTReader_MBuilding_构件设计_UC(Path, Num_Base + 1)
    '------------------------------------将首层墙柱最大轴压比及其构件编号写入general
        Sheets("g_M").Cells(8, 5) = Sheets("d_M").Cells(Num_Base + 3, 56)
        Sheets("g_M").Cells(8, 7) = Round(Sheets("d_M").Cells(Num_Base + 3, 57))
        Sheets("g_M").Cells(9, 5) = Sheets("d_M").Cells(Num_Base + 3, 58)
        Sheets("g_M").Cells(9, 7) = Round(Sheets("d_M").Cells(Num_Base + 3, 59))
    End If
    If CheckBox_WPJA Then
        Sheets("CR_M").Select
        ActiveWindow.Zoom = 70
        Sheets("WR_M").Select
        ActiveWindow.Zoom = 70
        For i = 1 To Num_all
            Call OUTReader_MBuilding_构件设计_UC(Path, i)
        Next
    '------------------------------------将全楼层墙柱最大轴压比及其构件编号写入general
        Sheets("g_M").Cells(8, 5) = Sheets("d_M").Cells(Num_Base + 3, 56)
        Sheets("g_M").Cells(8, 7) = Round(Sheets("d_M").Cells(Num_Base + 3, 57))
        Sheets("g_M").Cells(9, 5) = Sheets("d_M").Cells(Num_Base + 3, 58)
        Sheets("g_M").Cells(9, 7) = Round(Sheets("d_M").Cells(Num_Base + 3, 59))
    End If
 
    '----------------------------------------------------------写入wv02q.out文件数据
    If CheckBox_WV02Q Then
        Sheets("d_M").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_MBuilding_地震调整(Path)
    End If
    
'    '----------------------------------------------------------写入wdyna.out文件数据
'    If CheckBox_WDYNA Then
'        Call Addsh("Elastic_Dynamic_MBuilding")
'        Sheets("Elastic_Dynamic_MBuilding").Select
'        ActiveWindow.Zoom = 55
'        Call OUTReader_MBuilding_WDYNA(path)
'    End If

     Call gaoliang("M")

'Open "d:\dic.ini" For Output As #1
'        Print #1, Path
'    Close #1

    Sheets("d_M").Cells.NumberFormatLocal = "G/通用格式"
    Sheets("d_M").Columns("B:C").NumberFormatLocal = "0.00"
    Sheets("d_M").Columns("AT:AU").NumberFormatLocal = "0.00"
    Sheets("d_M").Columns("AH:AS").NumberFormatLocal = "0.000"
    Sheets("d_M").Columns("BB:BB").NumberFormatLocal = "0.0"

    Sheets("g_M").Select
    
End If




OUTReader_Main.Show 0

MsgBox "耗费时间: " & Timer - start

End Sub


'增加退出程序按钮
Sub QuitButton_Click()
    If Workbooks.Count > 1 Then
        ThisWorkbook.Close
    Else
        Application.Quit
    End If
End Sub


'窗体显示时通过按ESC键退出窗体，该Button控件已经被拉小隐藏退出后不能用F2调用，暂时冻结
'Private Sub CommandButtonP_Click()
'End
'End Sub

'反应谱数据绘图按钮
Private Sub Get_Figure_Click()

'计算运行时间
Dim start As Single
start = Timer

If CheckBox_PKPM Then

    Call OUTReader_Figure_Data("PKPM")
    
ElseIf CheckBox_YJK Then
    
    Call OUTReader_Figure_Data("YJK")
    
ElseIf CheckBox_MBuilding Then
    
    Call OUTReader_Figure_Data("MBuilding")
    
End If

    
OUTReader_Main.Show 0

MsgBox "耗费时间: " & Timer - start

End Sub

'时程曲线绘图按钮
Private Sub Get_Figure_Dyna_Click()

'计算运行时间
Dim start As Single
start = Timer

'Sheets("figure_dyna").Select

If CheckBox_PKPM Then

    Call OUTReader_Figure_Dyna("e_P")
  
ElseIf CheckBox_YJK Then
    
    Call OUTReader_Figure_Dyna("e_YJK")
        
End If

OUTReader_Main.Show 0

ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start


End Sub
'生成限值按钮
Private Sub LimitGe_Click()


If Height_TextBox.Text <= 150 Then
    DisLimit_TextBox.Text = DisLimit150_TextBox.Text
ElseIf Height_TextBox.Text >= 250 Then
    DisLimit_TextBox.Text = 500
Else
    DisLimit_TextBox.Text = 1 / (1 / DisLimit150_TextBox.Text + (Height_TextBox.Text - 150) / 100 * (1 / 500 - 1 / DisLimit150_TextBox.Text))
    DisLimit_TextBox.Text = Int(DisLimit_TextBox.Text)
End If

If Intensity_TextBox.Text = 6 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 0.8
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitX_TextBox.Text = 0.6
    Else
        RatioLimitX_TextBox.Text = 0.8 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (0.8 - 0.6)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 0.8
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitY_TextBox.Text = 0.6
    Else
        RatioLimitY_TextBox.Text = 0.8 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (0.8 - 0.6)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
ElseIf Intensity_TextBox.Text = 7 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 1.6
    ElseIf PeriodX_TextBox.Text > 5 Then
        'Debug.Print PeriodX_TextBox.Text
        RatioLimitX_TextBox.Text = 1.2
    Else
        RatioLimitX_TextBox.Text = 1.6 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (1.6 - 1.2)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 1.6
    ElseIf PeriodY_TextBox.Text > 5 Then
        Debug.Print PeriodX_TextBox.Text
        RatioLimitY_TextBox.Text = 1.2
    Else
        RatioLimitY_TextBox.Text = 1.6 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (1.6 - 1.2)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
ElseIf Intensity_TextBox.Text = 7.5 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 2.4
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitX_TextBox.Text = 1.8
    Else
        RatioLimitX_TextBox.Text = 2.4 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (2.4 - 1.8)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 2.4
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitY_TextBox.Text = 1.8
    Else
        RatioLimitY_TextBox.Text = 2.4 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (2.4 - 1.8)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
ElseIf Intensity_TextBox.Text = 8 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 3.2
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitX_TextBox.Text = 2.4
    Else
        RatioLimitX_TextBox.Text = 3.2 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (3.2 - 2.4)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 3.2
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitY_TextBox.Text = 2.4
    Else
        RatioLimitY_TextBox.Text = 3.2 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (3.2 - 2.4)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
ElseIf Intensity_TextBox.Text = 8.5 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 4.8
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitX_TextBox.Text = 3.6
    Else
        RatioLimitX_TextBox.Text = 4.8 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (4.8 - 3.6)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 4.8
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitY_TextBox.Text = 3.6
    Else
        RatioLimitY_TextBox.Text = 4.8 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (4.8 - 3.6)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
ElseIf Intensity_TextBox.Text = 9 Then
    If PeriodX_TextBox < 3.5 Then
        RatioLimitX_TextBox.Text = 6.4
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitX_TextBox.Text = 4.8
    Else
        RatioLimitX_TextBox.Text = 6.4 - (PeriodX_TextBox.Text - 3.5) / 1.5 * (6.4 - 4.8)
        RatioLimitX_TextBox.Text = Round(RatioLimitX_TextBox.Text, 2)
    End If
    
    If PeriodY_TextBox < 3.5 Then
        RatioLimitY_TextBox.Text = 6.4
    ElseIf PeriodX_TextBox.Text > 5 Then
        RatioLimitY_TextBox.Text = 4.8
    Else
        RatioLimitY_TextBox.Text = 6.4 - (PeriodY_TextBox.Text - 3.5) / 1.5 * (6.4 - 4.8)
        RatioLimitY_TextBox.Text = Round(RatioLimitY_TextBox.Text, 2)
    End If
Else
    MsgBox "请输入正确的设防烈度！"
End If


OUTReader_Main.Show 0

ThisWorkbook.Save


End Sub


'模型对比绘图按钮
Private Sub CommandButton4_Figure_Click()

'计算运行时间
Dim start As Single
start = Timer

If CheckBox4_PKPM And CheckBox4_MBuilding Then
    Call FigureCompare("d_P", "d_M", "SATWE", "Midas")
End If

If CheckBox4_PKPM And CheckBox4_YJK Then
    Call FigureCompare("d_P", "d_Y", "SATWE", "YJK")
End If

If CheckBox4_YJK And CheckBox4_MBuilding Then
    Call FigureCompare("d_Y", "d_M", "YJK", "Midas")
End If

If CheckBox4_ETABS And CheckBox4_MBuilding Then
    Call FigureCompare("d_E", "d_M", "ETABS", "Midas")
End If

If CheckBox4_ETABS And CheckBox4_PKPM Then
    Call FigureCompare("d_P", "d_E", "SATWE", "ETABS")
End If

If CheckBox4_ETABS And CheckBox4_YJK Then
    Call FigureCompare("d_E", "d_Y", "ETABS", "YJK")
End If
    


OUTReader_Main.Show 0

ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start


End Sub


'模型表格按钮
Private Sub CommandButton4_Data_Click()

'计算运行时间
Dim start As Single
start = Timer


If CheckBox4_PKPM And CheckBox4_MBuilding Then
    Call DataCompare("g_P", "g_M", "d_P", "d_M")
End If

If CheckBox4_PKPM And CheckBox4_YJK Then
    Call DataCompare("g_P", "g_Y", "d_P", "d_Y")
End If

If CheckBox4_YJK And CheckBox4_MBuilding Then
    Call DataCompare("g_Y", "g_M", "d_Y", "d_M")
End If

If CheckBox4_ETABS And CheckBox4_MBuilding Then
    Call DataCompare("g_E", "g_M", "d_E", "d_M")
End If

If CheckBox4_ETABS And CheckBox4_PKPM Then
    Call DataCompare("g_E", "g_P", "d_E", "d_P")
End If

If CheckBox4_ETABS And CheckBox4_YJK Then
    Call DataCompare("g_E", "g_Y", "d_E", "d_Y")
End If

    


OUTReader_Main.Show 0

ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start
End Sub

'模型表格按钮
Private Sub CommandButton4_FigureAll_Click()

'计算运行时间
Dim start As Single
start = Timer

Call CompareFigureAll


OUTReader_Main.Show 0

ThisWorkbook.Save
MsgBox "耗费时间: " & Timer - start
End Sub




'PKPM&YJK刚度比修正按钮
Private Sub modi_Click()

If CheckBox4_PKPM Then
    Call modi_stiff
    MsgBox "PKPM刚度及刚度比修正完成"
    Else
    Call modi_stiff_Y
    MsgBox "YJK刚度及刚度比修正完成"
End If



End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub dsheets_Click()
Dim i As Integer
For i = Worksheets.Count To 3 Step -1
    Application.DisplayAlerts = False
    Worksheets(i).Delete
    Application.DisplayAlerts = True
Next
End Sub

Private Sub 绘图限值_Click()

End Sub

Private Sub 结构信息_Click()

End Sub

'=======================================================================================================ETABS查看

Private Sub Get_ETABS_Path_Click()
'从MDB中读取数据
Dim fd As FileDialog
Application.FileDialog(msoFileDialogFilePicker).InitialFileName = ThisWorkbook.Path & "\"
Set fd = Application.FileDialog(msoFileDialogFilePicker)
fd.Filters.Clear
fd.Filters.Add "Access Databases", "*.mdb,*.accdb"
If fd.Show = -1 Then OUTReader_Main.TextBox_Path_ETABS = fd.SelectedItems(1)

'从MDB中读取数据
'If Option_ETABS_MDB Then
'    Dim WBToWrite As String
'    WBToWrite = Application.GetOpenFilename("Access Files (*.mdb), *.mdb")
'    If WBToWrite <> "False" Then
'        OUTReader_Main.TextBox_Path_ETABS = WBToWrite
'    End If
'End If

End Sub


Private Sub ETABS_Read_LOAD_Click()
'读取ETABS荷载工况名，赋工况名到相应复选框
Dim MDB_Path As String

MDB_Path = OUTReader_Main.TextBox_Path_ETABS.Text

Call Etabs_Load_Case(MDB_Path)

End Sub

Private Sub Read_ETABS_Data_Click()
'读取ETABS数据
Dim MDB_Path As String
MDB_Path = OUTReader_Main.TextBox_Path_ETABS.Text

'定义表格
Call Addsh("g_E")
Call Addsh("d_E")
Call AddHeadline("g_E", "d_E")
Sheets("g_E").Cells(3, 4) = MDB_Path

Call ETABS_DATA_READ(MDB_Path)
Call ETABS_HIST_DATA(MDB_Path)

Call gaoliang("E")

End Sub
Private Sub ETABSMOB_Click()
'读取ETABS位移角剪重比修正

Call ETABS_DATA_CALC("ETABSMOB")
'Debug.Print "aa"

End Sub


Private Sub Fig_ETABS_Click()
'ETABS反应谱数据画图
Call OUTReader_Figure_Data("ETABS")

End Sub


Private Sub ETABS_HISTFIG_Click()
'ETABS时程分析数据画图
Call ETABS_HIST_Fig("e_E")
End Sub


