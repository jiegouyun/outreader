Attribute VB_Name = "模块_高亮"
'设置Range条件格式的代码

'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2015/4/19
'更新内容：
'1.添加ETABS

'////////////////////////////////////////////////////////////////////////////////////////////
'更新时间:2014/4/18

'更新内容：
'1.添加判断，无数据则不高亮

Sub gaoliang(soft As String)
On Error Resume Next
'-------------------------------------------------------------------------------------------高亮最值
Dim sht As String
If soft = "P" Then
    sht = "d_P"
ElseIf soft = "Y" Then
    sht = "d_Y"
ElseIf soft = "M" Then
    sht = "d_M"
ElseIf soft = "E" Then
    sht = "d_E"
End If


Num_all = Sheets(sht).range("a65536").End(xlUp)
Debug.Print "总楼层="; Num_all

Dim i_RowID As Integer
Dim i_Rng As range

'---------------------------------------------------------刚度比
For ii = 2 To 3
If Worksheets(sht).Cells(3, ii) <> "" Then
    Dim R As range
    Set R = Worksheets(sht).range(Worksheets(sht).Cells(3, ii), Worksheets(sht).Cells(Num_all + 1, ii))
    Call maxormin(R, "min", sht & "!R3C" & CStr(ii) & ":R" & CStr(Num_all + 1) & "C" & CStr(ii))
    End If
Next

'---------------------------------------------------------承载力比
For ii = 46 To 47
If Worksheets(sht).Cells(3, ii) <> "" Then
    Set R = Worksheets(sht).range(Worksheets(sht).Cells(3, ii), Worksheets(sht).Cells(Num_all + 1, ii))
    Call maxormin(R, "min", sht & "!R3C" & CStr(ii) & ":R" & CStr(Num_all + 1) & "C" & CStr(ii))
End If
Next

'---------------------------------------------------------质量比
ii = 55
    If Worksheets(sht).Cells(3, ii) <> "" Then
    Set R = Worksheets(sht).range(Worksheets(sht).Cells(4, ii), Worksheets(sht).Cells(Num_all + 2, ii))
    Call maxormin(R, "max", sht & "!R4C" & CStr(ii) & ":R" & CStr(Num_all + 2) & "C" & CStr(ii))
End If


'---------------------------------------------------------位移角
For ii = 26 To 33
If Worksheets(sht).Cells(3, ii) <> "" Then
    Set R = Worksheets(sht).range(Worksheets(sht).Cells(3, ii), Worksheets(sht).Cells(Num_all + 2, ii))
    Call maxormin(R, "min", sht & "!R3C" & CStr(ii) & ":R" & CStr(Num_all + 2) & "C" & CStr(ii))
End If
Next

'---------------------------------------------------------位移比
For ii = 34 To 45
If Worksheets(sht).Cells(3, ii) <> "" Then
    Set R = Worksheets(sht).range(Worksheets(sht).Cells(3, ii), Worksheets(sht).Cells(Num_all + 2, ii))
    Call maxormin(R, "max", sht & "!R3C" & CStr(ii) & ":R" & CStr(Num_all + 2) & "C" & CStr(ii))
End If
Next

End Sub
