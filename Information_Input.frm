VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Information_Input 
   Caption         =   "Input"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7515
   OleObjectBlob   =   "Information_Input.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "Information_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/2

'更新内容：
'1.添加默认值

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/30 20:56

'更新内容：
'1.添加MBuilding模块，但是因为其位移比模块读取最大位移比有问题，故先锁定，待解决该问题可解锁

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/7/23 13:56

'更新内容：
'1.添加调试模式中手动输入内容界面



'////////////////////////////////////////////////////////////////////////////

'点击取消，退出Information_Input窗体，打开OUTRader_Main窗体
Private Sub CommandButton_cancel_Click()
Information_Input.Hide
OUTReader_Main.Show 0
End Sub

'点击确定，开始读取数据
Private Sub CommandButton_OK_Click()
Information_Input.TextBox_Folder.Text = OUTReader_Main.TextBox_Path_2.Text
Information_Input.Hide
If OUTReader_Main.OptionButton_Excel Then
Call Test_XX(2)
End If

'从OUT数据中读取数据
If OUTReader_Main.OptionButton_OUT Then
Call Debug_Mode(2)
'OUTReader_Main.TextBox_Path_2.Text = OUTReader_Main.TextBox_Path_2.Text & "\" & ThisWorkbook.name
'OUTReader_Main.TextBox_Path_2.Text = ThisWorkbook.Path
Call Test_XX(2)
End If

'Dim WB_New_Full As String
'Dim WB_Orig_Full As String
'WB_Orig_Full = OUTReader_Main.TextBox_Path_2.Text
'WB_New_Full = OUTReader_Main.TextBox_Path_3.Text
'Path = Workbooks(WB_Orig).Worksheets("debug").Cells(2, 2)
'Debug.Print Path
'Call Data_Summarization(WB_Orig_Full, WB_New_Full)
End Sub

Public Sub Debug_Mode(X As Integer)
X = 1

'计算运行时间
Dim start As Single
start = Timer

ThisWorkbook.Activate
OUTReader_Main.Hide

'==========================================================================================定义主要辅助变量

Dim Path As String
'Dim General, Distribution, Program As String


'--------------------------------------------------------------删除原有数据
'Dim i_s As Integer
'For i_s = Sheets.Count To 1 Step -1
    'Sheets(i_s).Cells.Clear
'Next


'--------------------------------------------------------------获取数据存储路径
Path = OUTReader_Main.TextBox_Path_2.Text


'==========================================================================================提取PKPM数据
If OUTReader_Main.CheckBox_PKPM_2 Then

    '----------------------------------------------------------生成表格表头
    'Program = "PKPM"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_P")
    Call Addsh("d_P")
    Call AddHeadline("g_P", "d_P")

    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)

    'OptionButton_WPJF.Locked = True
    'OptionButton_WPJA.Locked = True
    Sheets("d_P").Select
    ActiveWindow.Zoom = 55
    Call OUTReader_PKPM_WMASS(Path)

    '----------------------------------------------------------写入wzq.out文件数据
    Sheets("d_P").Select
    ActiveWindow.Zoom = 55
    Call OUTReader_PKPM_WZQ(Path)

    '----------------------------------------------------------写入wdisp.out文件数据
    Sheets("d_P").Select
    Call OUTReader_PKPM_WDISP(Path)
    Sheets("g_P").Select

End If

'==========================================================================================提取YJK数据
If OUTReader_Main.CheckBox_YJK_2 Then

    '----------------------------------------------------------生成表格表头
    'Program = "YJK"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_Y")
    Call Addsh("d_Y")
    Call AddHeadline("g_Y", "d_Y")

    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)

    'OptionButton_WPJF.Locked = True
    'OptionButton_WPJA.Locked = True
    Sheets("d_Y").Select
    ActiveWindow.Zoom = 55
    Call OUTReader_YJK_WMASS(Path)

    '----------------------------------------------------------写入wzq.out文件数据
    Sheets("d_Y").Select
    ActiveWindow.Zoom = 55
    Call OUTReader_YJK_WZQ(Path)
        
    '----------------------------------------------------------写入wdisp.out文件数据
    Sheets("d_Y").Select
    Call OUTReader_YJK_WDISP(Path)
    Sheets("g_Y").Select
            
End If

'==========================================================================================提取MBuilding数据
If OUTReader_Main.CheckBox_MBuilding_2 Then

    '----------------------------------------------------------生成表格表头
    'Program = "MBuilding"
    'General = "General_" & Program
    'Distribution = "Distribution_" & Program
    'Call Addsh(General)
   ' Call Addsh(Distribution)
    Call Addsh("g_M")
    Call Addsh("d_M")
    Call AddHeadline("g_M", "d_M")

    '----------------------------------------------------------写入wmass.out文件数据(默认必须选择)
        'OptionButton_WPJF.Locked = True
        'OptionButton_WPJA.Locked = True
        Sheets("d_M").Select
        ActiveWindow.Zoom = 55
        Call OUTReader_MBuilding_总信息(Path)
        Call OUTReader_MBuilding_侧向刚度(Path)
        Call OUTReader_MBuilding_抗剪承载力(Path)

    '----------------------------------------------------------写入wzq.out文件数据
    Sheets("d_M").Select
    ActiveWindow.Zoom = 55
    Call OUTReader_MBuilding_周期振型(Path)
        
    '----------------------------------------------------------写入wdisp.out文件数据
    Sheets("d_M").Select
    Call OUTReader_MBuilding_结构位移(Path)
    Sheets("g_M").Select
            
End If


OUTReader_Main.Show 0

MsgBox "耗费时间: " & Timer - start

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox_Operate_Change()

End Sub

Private Sub TextBox_Stage_Change()

End Sub

Private Sub UserForm_Click()

End Sub

