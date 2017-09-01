Attribute VB_Name = "模块_修改图标格式代码"
Option Explicit

Sub ChangeDataName()
Attribute ChangeDataName.VB_ProcData.VB_Invoke_Func = "g\n14"
'
'
Dim name1 As String, name2 As String

name1 = InputBox("请输入数据一的名称", "更改数据名称", "方案一")
name2 = InputBox("请输入数据二的名称", "更改数据名称", "方案二")

    ActiveChart.SeriesCollection(1).Select
    Selection.name = name1
    ActiveChart.SeriesCollection(2).Select
    Selection.name = name2
    
End Sub



Option Explicit

Sub changeformat()
'
' 修改图表的格式
' 仅适用于对比的图表
' 系列1改为绿色，系列2改为紫色，背景填充改为白色

Dim i As Integer

For i = 1 To ActiveSheet.ChartObjects.Count
ActiveSheet.ChartObjects("图表 " & i).Activate
   ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
    End With
    ActiveChart.PlotArea.Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
Next

End Sub

