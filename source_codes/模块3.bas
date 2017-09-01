Attribute VB_Name = "模块3"

Option Explicit

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/13
'1.图表高宽改为外部参数输入，方便修改


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/19
'1.



Sub CompareFigureAll()

'----------------------------------------------------------定义数组，判断选择软件的类型
Dim soft(3)
Dim dis(3)
Dim names(3)

If OUTReader_Main.CheckBox4_PKPM Then
  soft(0) = 1
Else
  soft(0) = 0
End If

If OUTReader_Main.CheckBox4_YJK Then
  soft(1) = 1
Else
  soft(1) = 0
End If

If OUTReader_Main.CheckBox4_MBuilding Then
  soft(2) = 1
Else
  soft(2) = 0
End If

If OUTReader_Main.CheckBox4_ETABS Then
  soft(3) = 1
Else
  soft(3) = 0
End If

dis(0) = "d_P"
dis(1) = "d_Y"
dis(2) = "d_M"
dis(3) = "d_E"

names(0) = "SATWE"
names(1) = "YJK"
names(2) = "Midas Building"
names(3) = "ETABS"


'Debug.Print dis(3)

'选择的软件个数an
Dim an As Integer
an = soft(0) + soft(1) + soft(2) + soft(3)

'Debug.Print an
Dim figuresofts(), figurenames()
ReDim Preserve figuresofts(an - 1)
ReDim Preserve figurenames(an - 1)

Dim n As Integer
Dim n1 As Integer
n1 = 0
For n = 0 To 3
    If soft(n) = 1 Then
        figuresofts(n1) = dis(n)
        figurenames(n1) = names(n)
        n1 = n1 + 1
    End If
Next

'----------------------------------------------------------确定楼层数

Dim ela As String
For n = 0 To 3
    If soft(n) = 1 Then
        ela = dis(n)
    End If
Next
'Debug.Print ela
'确定楼层数
Sheets(ela).Select
Dim NN As Integer: NN = Cells(Rows.Count, "j").End(3).Row - 2
'Debug.Print "NN=" & NN
'定义全局变量楼层总数，确保在没有反应谱数据时时程画图楼层正确
Num_all = NN
  
  

'----------------------------------------------------------删除已有figure工作表
Dim sh As Worksheet

'搜寻已有的工作表的名称
For Each sh In Worksheets
    '如果与新定义的工作表名相同，则删除工作表
    If sh.name = "F_All" Then
        sh.Delete
    End If
Next

'新建一个工作表
With Worksheets
    Set sh = .Add(After:=Worksheets(.Count))
    sh.name = "F_All"
    End With


'----------------------------------------------------------迁移处理位移角数据

For n = 0 To an - 1
    With Sheets(dis(n)).range("BI3:" & "BP" & Num_all + 2)
        .FormulaR1C1 = "=1/" & dis(n) & "!RC[-35]"
        .Font.ColorIndex = 1
        .Locked = True
    End With
Next





'----------------------------------------------------------绘图

Dim i As Integer
Sheets("F_All").Select

'绘图公用部分
Dim sheetname(), range_X(), range_Y(0), name_S(), name_XY(1)
ReDim Preserve sheetname(an)
ReDim Preserve range_X(an - 1)
ReDim Preserve name_S(an - 1)
Dim Location(1) As Integer

'图表高宽
Dim Width As Integer, Hight As Integer
Width = 207
Hight = 284


For i = 0 To an - 1
    sheetname(i) = figuresofts(i)
Next i
sheetname(an) = "F_All"

range_Y(0) = "R3C1:R" & NN + 2 & "C1"
name_XY(1) = "层数"

For i = 0 To an - 1
    name_S(i) = figurenames(i)
Next i



'X向刚度比
For i = 0 To an - 1
    range_X(i) = "R3C2:R" & NN + 2 & "C2"
Next i
name_XY(0) = "X向刚度比"
Location(0) = 0 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向刚度
For i = 0 To an - 1
    range_X(i) = "R3C4:R" & NN + 2 & "C4"
Next i
name_XY(0) = "X向刚度"
Location(0) = 1 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下X向剪力(kN)
For i = 0 To an - 1
    range_X(i) = "R3C6:R" & NN + 2 & "C6"
Next i
name_XY(0) = "风荷载下X向剪力(kN)"
Location(0) = 2 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下Y向剪力(kN)
For i = 0 To an - 1
    range_X(i) = "R3C8:R" & NN + 2 & "C8"
Next i
name_XY(0) = "风荷载下Y向剪力(kN)"
Location(0) = 3 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下X向剪力(kN)
For i = 0 To an - 1
    range_X(i) = "R3C10:R" & NN + 2 & "C10"
Next i
name_XY(0) = "地震作用下X向剪力(kN)"
Location(0) = 4 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下Y向剪力(kN)
For i = 0 To an - 1
    range_X(i) = "R3C14:R" & NN + 2 & "C14"
Next i
name_XY(0) = "地震作用下Y向剪力(kN)"
Location(0) = 5 * Width
Location(1) = 0 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向刚度比
For i = 0 To an - 1
    range_X(i) = "R3C3:R" & NN + 2 & "C3"
Next i
name_XY(0) = "Y向刚度比"
Location(0) = 0 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向刚度
For i = 0 To an - 1
    range_X(i) = "R3C5:R" & NN + 2 & "C5"
Next i
name_XY(0) = "Y向刚度"
Location(0) = 1 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下X向弯矩(kNm)
For i = 0 To an - 1
    range_X(i) = "R3C7:R" & NN + 2 & "C7"
Next i
name_XY(0) = "风荷载下X向弯矩(kNm)"
Location(0) = 2 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'风荷载下Y向弯矩(kNm)
For i = 0 To an - 1
    range_X(i) = "R3C9:R" & NN + 2 & "C9"
Next i
name_XY(0) = "风荷载下Y向弯矩(kNm)"
Location(0) = 3 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下X向弯矩(kNm)
For i = 0 To an - 1
    range_X(i) = "R3C11:R" & NN + 2 & "C11"
Next i
name_XY(0) = "地震作用下X向弯矩(kNm)"
Location(0) = 4 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'地震作用下Y向弯矩(kNm)
For i = 0 To an - 1
    range_X(i) = "R3C15:R" & NN + 2 & "C15"
Next i
name_XY(0) = "地震作用下Y向弯矩(kNm)"
Location(0) = 5 * Width
Location(1) = 1 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX工况下位移角
For i = 0 To an - 1
    range_X(i) = "R3C61:R" & NN + 2 & "C61"
Next i
name_XY(0) = "EX工况下位移角"
Location(0) = 0 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'EY工况下位移角
For i = 0 To an - 1
    range_X(i) = "R3C65:R" & NN + 2 & "C65"
Next i
name_XY(0) = "EY工况下位移角"
Location(0) = 1 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'WX工况下位移角
For i = 0 To an - 1
    range_X(i) = "R3C64:R" & NN + 2 & "C64"
Next i
name_XY(0) = "WX工况下位移角"
Location(0) = 2 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'WY工况下位移角
For i = 0 To an - 1
    range_X(i) = "R3C68:R" & NN + 2 & "C68"
Next i
name_XY(0) = "WY工况下位移角"
Location(0) = 3 * Width
Location(1) = 2 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight, "#/###0")

'EX+工况下位移比
For i = 0 To an - 1
    range_X(i) = "R3C35:R" & NN + 2 & "C35"
Next i
name_XY(0) = "EX+工况下位移比"
Location(0) = 0 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX-工况下位移比
For i = 0 To an - 1
    range_X(i) = "R3C36:R" & NN + 2 & "C36"
Next i
name_XY(0) = "EX-工况下位移比"
Location(0) = 1 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY+工况下位移比
For i = 0 To an - 1
    range_X(i) = "R3C38:R" & NN + 2 & "C38"
Next i
name_XY(0) = "EY+工况下位移比"
Location(0) = 2 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY-工况下位移比
For i = 0 To an - 1
    range_X(i) = "R3C39:R" & NN + 2 & "C39"
Next i
name_XY(0) = "EY-工况下位移比"
Location(0) = 3 * Width
Location(1) = 3 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX+工况下层间位移比
For i = 0 To an - 1
    range_X(i) = "R3C41:R" & NN + 2 & "C41"
Next i
name_XY(0) = "EX+工况下层间位移比"
Location(0) = 0 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EX-工况下层间位移比
For i = 0 To an - 1
    range_X(i) = "R3C42:R" & NN + 2 & "C42"
Next i
name_XY(0) = "EX-工况下层间位移比"
Location(0) = 1 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY+工况下层间位移比
For i = 0 To an - 1
    range_X(i) = "R3C44:R" & NN + 2 & "C44"
Next i
name_XY(0) = "EY+工况下层间位移比"
Location(0) = 2 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'EY-工况下层间位移比
For i = 0 To an - 1
    range_X(i) = "R3C45:R" & NN + 2 & "C45"
Next i
name_XY(0) = "EY-工况下层间位移比"
Location(0) = 3 * Width
Location(1) = 4 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向剪重比
For i = 0 To an - 1
    range_X(i) = "R3C12:R" & NN + 2 & "C12"
Next i
name_XY(0) = "X向剪重比"
Location(0) = 0 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向剪重比
For i = 0 To an - 1
    range_X(i) = "R3C16:R" & NN + 2 & "C16"
Next i
name_XY(0) = "Y向剪重比"
Location(0) = 1 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'X向抗剪承载力比
For i = 0 To an - 1
    range_X(i) = "R3C46:R" & NN + 2 & "C46"
Next i
name_XY(0) = "X向抗剪承载力比"
Location(0) = 2 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'Y向抗剪承载力比
For i = 0 To an - 1
    range_X(i) = "R3C47:R" & NN + 2 & "C47"
Next i
name_XY(0) = "Y向抗剪承载力比"
Location(0) = 3 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'单位面积质量
For i = 0 To an - 1
    range_X(i) = "R3C54:R" & NN + 2 & "C54"
Next i
name_XY(0) = "单位面积质量"
Location(0) = 4 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'质量比
For i = 0 To an - 1
    range_X(i) = "R3C55:R" & NN + 2 & "C55"
Next i
name_XY(0) = "质量比"
Location(0) = 5 * Width
Location(1) = 5 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱X向地震剪力百分比
For i = 0 To an - 1
    range_X(i) = "R3C49:R" & NN + 2 & "C49"
Next i
name_XY(0) = "框架柱X向地震剪力百分比"
Location(0) = 0 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱Y向地震剪力百分比
For i = 0 To an - 1
    range_X(i) = "R3C52:R" & NN + 2 & "C52"
Next i
name_XY(0) = "框架柱Y向地震剪力百分比"
Location(0) = 1 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱X向地震剪力调整系数
For i = 0 To an - 1
    range_X(i) = "R3C50:R" & NN + 2 & "C50"
Next i
name_XY(0) = "框架柱X向地震剪力调整系数"
Location(0) = 2 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)

'框架柱Y向地震剪力调整系数
For i = 0 To an - 1
    range_X(i) = "R3C53:R" & NN + 2 & "C53"
Next i
name_XY(0) = "框架柱Y向地震剪力调整系数"
Location(0) = 3 * Width
Location(1) = 6 * Hight
Call add_chart_array(sheetname(), an, range_X(), range_Y(), name_S(), name_XY(), Location(), Width, Hight)


End Sub



