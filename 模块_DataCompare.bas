Attribute VB_Name = "模块_DataCompare"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/24

'更新内容：
'1.修改d表对比缺少楼层的bug；
'2.分离表名，如PKPM和YJK对比，将表明定为：g_P&Y和d_P&Y

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/11/15

'更新内容：
'1.修改了一个 写质量比 时范围

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2013/8/24

'更新内容：
'1.将dis表也进行对比



Sub DataCompare(model_1_g, model_2_g, model_1_d, model_2_d)

Dim g, d As String
'g = "g_compare"
'd = "d_compare"

If model_1_g = "g_P" And model_2_g = "g_M" Then
    g = "g_P&M"
    d = "d_P&M"
End If
If model_1_g = "g_P" And model_2_g = "g_Y" Then
    g = "g_P&Y"
    d = "d_P&Y"
End If
If model_1_g = "g_Y" And model_2_g = "g_M" Then
    g = "g_Y&M"
    d = "d_Y&M"
End If
If model_1_g = "g_E" And model_2_g = "g_M" Then
    g = "g_E&M"
    d = "d_E&M"
End If
If model_1_g = "g_E" And model_2_g = "g_P" Then
    g = "g_E&P"
    d = "d_E&P"
End If
If model_1_g = "g_E" And model_2_g = "g_Y" Then
    g = "g_E&Y"
    d = "d_E&Y"
End If




'If CheckBox4_PKPM And CheckBox4_MBuilding Then
'    Call DataCompare("g_P", "g_M", "d_P", "d_M")
'End If
'
'If CheckBox4_PKPM And CheckBox4_YJK Then
'    Call DataCompare("g_P", "g_Y", "d_P", "d_Y")
'End If
'
'If CheckBox4_YJK And CheckBox4_MBuilding Then
'    Call DataCompare("g_Y", "g_M", "d_Y", "d_M")
'End If
'
'If CheckBox4_ETABS And CheckBox4_MBuilding Then
'    Call DataCompare("g_E", "g_M", "d_E", "d_M")
'End If
'
'If CheckBox4_ETABS And CheckBox4_PKPM Then
'    Call DataCompare("g_E", "g_P", "d_E", "d_P")
'End If
'
'If CheckBox4_ETABS And CheckBox4_YJK Then
'    Call DataCompare("g_E", "g_Y", "d_E", "d_Y")
'End If

'--------------------------------------------------------------------------------创建工作表
Call Addsh(g)
Call Addsh(d)
Call AddHeadline(g, d)

'----------------------------------------------------------读取楼层总数
Num_all = Sheets(model_1_d).[A65536].End(xlUp).Row - 2
Debug.Print Num_all

'--------------------------------------------------------------------------------将对比数据写入工作表
Sheets(g).Select

Dim ii, jj As Integer

For ii = 3 To 7
    For jj = 4 To 7
        If Sheets(model_1_g).Cells(ii, jj).Text <> Sheets(model_2_g).Cells(ii, jj).Text Then
         Debug.Print "a"
            Sheets(g).Cells(ii, jj).Value = Format(Sheets(model_1_g).Cells(ii, jj).Value, "0.00") & " | " & Format(Sheets(model_2_g).Cells(ii, jj).Value, "0.00")
            Sheets(g).Cells(ii, jj).Font.Size = 9
        Else
            Sheets(g).Cells(ii, jj).Value = Sheets(model_1_g).Cells(ii, jj).Value
        End If
    Next
Next

For ii = 8 To 19
    For jj = 4 To 7
        If Sheets(model_1_g).Cells(ii, jj).Text <> Sheets(model_2_g).Cells(ii, jj).Text Then
         Debug.Print "a"
            Sheets(g).Cells(ii, jj).Value = Sheets(model_1_g).Cells(ii, jj).Text & " | " & Sheets(model_2_g).Cells(ii, jj).Text
            Sheets(g).Cells(ii, jj).Font.Size = 9
        Else
            Sheets(g).Cells(ii, jj).Value = Sheets(model_1_g).Cells(ii, jj).Value
        End If
    Next
Next

For ii = 20 To 51
    For jj = 4 To 7
        If Sheets(model_1_g).Cells(ii, jj).Text <> Sheets(model_2_g).Cells(ii, jj).Text Then
         Debug.Print "a"
            Sheets(g).Cells(ii, jj).Value = Format(Sheets(model_1_g).Cells(ii, jj).Value, "0.00") & " | " & Format(Sheets(model_2_g).Cells(ii, jj).Value, "0.00")
            Sheets(g).Cells(ii, jj).Font.Size = 9
        Else
            Sheets(g).Cells(ii, jj).Value = Sheets(model_1_g).Cells(ii, jj).Value
        End If
    Next
Next


Sheets(d).Select

range(Sheets(d).Cells(3, 1), Sheets(d).Cells(Num_all + 2, 59)).Cells.Font.Size = "8"

'---------------------------------------写层号
For ii = 3 To Num_all + 2
        Sheets(d).Cells(ii, 1).Value = Sheets(model_1_d).Cells(ii, 1).Value
Next
'---------------------------------------写刚度比
For ii = 3 To Num_all + 2
    For jj = 2 To 3
        Sheets(d).Cells(ii, jj).Value = Format(Sheets(model_1_d).Cells(ii, jj).Value, "0.00") & " | " & Format(Sheets(model_2_d).Cells(ii, jj).Value, "0.00")
    Next
Next
'---------------------------------------写刚度比承载力比
For ii = 3 To Num_all + 2
    For jj = 4 To 45
        Sheets(d).Cells(ii, jj).Value = Sheets(model_1_d).Cells(ii, jj).Value & " | " & Sheets(model_2_d).Cells(ii, jj).Value
    Next
Next

'---------------------------------------写承载力比、质量比
For ii = 3 To Num_all + 2
    For jj = 46 To 53
        Sheets(d).Cells(ii, jj).Value = Format(Sheets(model_1_d).Cells(ii, jj).Value, "0.00") & " | " & Format(Sheets(model_2_d).Cells(ii, jj).Value, "0.00")
    Next
Next
'---------------------------------------写质量比
For ii = 3 To Num_all + 2
    For jj = 54 To 59
        Sheets(d).Cells(ii, jj).Value = Sheets(model_1_d).Cells(ii, jj).Value & " | " & Sheets(model_2_d).Cells(ii, jj).Value
    Next
Next

Sheets(d).Cells.EntireColumn.AutoFit

Sheets(g).Select

End Sub
