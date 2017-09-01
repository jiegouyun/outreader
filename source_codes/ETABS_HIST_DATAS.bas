Attribute VB_Name = "ETABS_HIST_DATAS"
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                      ETABS时程分析数据读取                           ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/3/11
'更新内容:
'1.修正汇总表中65%，135%，80%，120%反应谱数据位置错误

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/5/12
'更新内容:
'1.修正X向与Y向选择时程波数量不同时Y向平均值计算错误

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/11
'更新内容:
'1.使用excel自带函数match求最值所在行，避免with报错
'2.更正剪力和弯矩先取绝对值再取大值，更新temp_a等变量类型为string,原int类型有误


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/7
'更新内容:
'1.时程工况数据提取单列出来了



Public Sub ETABS_HIST_DATA(MDB_Path As String)

'MDB_Path = OUTReader_Main.TextBox_Path_ETABS.Text
'Num_all = 20

'计算X向何Y向时程函数个数
Dim Num_X, Num_Y, N_LX, N_LY As Integer
Dim name_X() As String
Dim name_Y() As String

Num_X = 0
Num_Y = 0

N_LX = OUTReader_Main.ListBox_TH_X.ListCount
N_LY = OUTReader_Main.ListBox_TH_Y.ListCount

If N_LX + N_LY = 0 Then Exit Sub

ReDim name_X(N_LX - 1)
ReDim name_Y(N_LY - 1)

For i = 0 To OUTReader_Main.ListBox_TH_X.ListCount - 1
    If OUTReader_Main.ListBox_TH_X.Selected(i) = True Then
        name_X(Num_X) = OUTReader_Main.ListBox_TH_X.List(i)
        Num_X = Num_X + 1
        'Debug.Print Name_X(i)
    End If
Next i

For i = 0 To OUTReader_Main.ListBox_TH_Y.ListCount - 1
    If OUTReader_Main.ListBox_TH_Y.Selected(i) = True Then
        name_Y(Num_Y) = OUTReader_Main.ListBox_TH_Y.List(i)
        Num_Y = Num_Y + 1
        'Debug.Print Name_Y(i)
    End If
Next i

If Num_X + Num_Y = 0 Then Exit Sub

'写入时程表格
'======================================================================================================提示反应谱数据
'If Sheets("d_P").Cells(3, 6) = "" Or Sheets("d_P").Cells(3, 10) = "" Or Sheets("d_P").Cells(3, 18) = "" Then
   ' MsgBox "缺少反应谱数据，请返回重新选择。"
'Else


'======================================================================================================设定表格Elastic-Dynamic的格式
'定义结果表格名称
Dim ela As String
ela = "e_E"

Call Addsh(ela)
Sheets(ela).Select
ActiveWindow.Zoom = 55

'清除工作表所有内容
Sheets(ela).Cells.Clear

'======================================================================================================添加表格Elastic-Dynamic的标题

For i = 1 To Num_all
            
    Sheets(ela).Cells(i + 2, 9) = i
                
Next i

'------------------------------------------------------工作表Elastic-Dynamic内的标题格式
With Sheets(ela)
    
    '设置字体
    .Cells.Font.name = "Times New Roman"
    '设置字体大小
    .Cells.Font.Size = "11"
    '水平居中
    .Cells.HorizontalAlignment = xlCenter
    '垂直居中
    .Cells.VerticalAlignment = xlCenter
    '不自动换行
    .Cells.WrapText = False
    
    '-------------------------------------------------汇总表格区1
    
    '项目信息区标题
    .Cells(2, 1) = "时程波总数"
    .Cells(2, 2) = Num_X + Num_Y
    .Cells(2, 3) = "X向"
    .Cells(2, 4) = Num_X
    .Cells(2, 5) = "Y向"
    .Cells(2, 6) = Num_Y
    .Cells(4, 1) = "作用工况"
    .Cells(4, 2) = "作用方向=0°"
    .Cells(4, 5) = "作用方向=90°"
    .Cells(5, 2) = "基底剪力"
    .Cells(5, 3) = "时程/反应谱"
    .Cells(5, 4) = "位移角"
    .Cells(5, 5) = "基底剪力"
    .Cells(5, 6) = "时程/反应谱"
    .Cells(5, 7) = "位移角"
    .range("A4:A5").MergeCells = True
    .range("B4:D4").MergeCells = True
    .range("E4:G4").MergeCells = True
    
    '-------------------------------------------------汇总表格区2
    
    '项目信息区标题
    '.Cells(18, 1) = "作用工况"
    '.Cells(18, 2) = "作用方向=0°"
    '.Cells(18, 5) = "作用方向=90°"
    '.Cells(19, 2) = "位移角"
    '.Cells(19, 3) = "时程/反应谱"
    '.Cells(19, 4) = "平均值/反应谱"
    '.Cells(19, 5) = "基底剪力"
    '.Cells(19, 6) = "时程/反应谱"
    '.Cells(19, 7) = "平均值/反应谱"
    '.Cells(20, 1) = "反应谱"
    '.range("A18:A19").MergeCells = True
    '.range("B18:D18").MergeCells = True
    '.range("E18:G18").MergeCells = True
    
End With

'加表格线
Call AddFormLine(ela, "A1:DZ200")

'=====================================================================================================将名称写入汇总表
Dim m As Integer
Dim Temp_Colour As Integer
Temp_Colour = -1

For m = 1 To Num_X
    Sheets(ela).Cells(m + 5, 1) = name_X(m - 1)
    
    '将标题项写入层分布表
    Sheets(ela).range(Cells(1, (m - 1) * 3 + 10), Cells(1, (m - 1) * 3 + 12)).MergeCells = True
    Sheets(ela).Cells(1, 10 + (m - 1) * 3) = name_X(m - 1)
    Sheets(ela).Cells(2, (m - 1) * 3 + 10) = "层间位移角X"
    Sheets(ela).Cells(2, (m - 1) * 3 + 11) = "剪力X"
    Sheets(ela).Cells(2, (m - 1) * 3 + 12) = "倾覆弯矩X"
    
    
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    
    Sheets(ela).range(Cells(1, (m - 1) * 3 + 10), Cells(2, (m - 1) * 3 + 12)).Interior.color = Colour
    Temp_Colour = -1 * Temp_Colour

Next m

For m = 1 To Num_Y
    Sheets(ela).Cells(m + Num_X + 5, 1) = name_Y(m - 1)
    
    '将标题项写入层分布表
    Sheets(ela).range(Cells(1, (m - 1) * 3 + 10 + 3 * Num_X), Cells(1, (m - 1) * 3 + 12 + 3 * Num_X)).MergeCells = True
    Sheets(ela).Cells(1, 10 + 3 * Num_X + (m - 1) * 3) = name_Y(m - 1)
    Sheets(ela).Cells(2, (m - 1) * 3 + 10 + 3 * Num_X) = "层间位移角Y"
    Sheets(ela).Cells(2, (m - 1) * 3 + 11 + 3 * Num_X) = "剪力Y"
    Sheets(ela).Cells(2, (m - 1) * 3 + 12 + 3 * Num_X) = "倾覆弯矩Y"
    
    
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    
    Sheets(ela).range(Cells(1, (m - 1) * 3 + 10 + 3 * Num_X), Cells(2, (m - 1) * 3 + 12 + 3 * Num_X)).Interior.color = Colour
    Temp_Colour = -1 * Temp_Colour

Next m



'************************************************************************************************************************读取时程数据

'定义变量
Dim E_Connect As ADODB.Connection
Dim E_RcdSet1 As ADODB.Recordset
Dim E_RcdSet2 As ADODB.Recordset
Dim E_RxSchema As ADODB.Recordset

Dim StrSQL1 As String
Dim StrSQL2 As String

Dim Temp_a, Temp_b, Temp_a1, Temp_b1 As String

'判断Access文件是否存在
If Dir(MDB_Path) = " " Then
  MsgBox "MDB文件不存在！请核实！", vbExclamation, "无法连接数据库"
  Exit Sub
End If

'使用ADO连接Access文件
'对于Access 2007 及高版本EXCEL
Set E_Connect = New ADODB.Connection
E_Connect.Open ConnectionString:="Provider=Microsoft.Ace.OLEDB.12.0;" & "Data Source =" & MDB_Path & ";" '& "Extended Properties=Excel 12.0;"
'对于早期版本的Access和Excel使用
'myConnect.Open ConnectionString:="Provider=Microsoft.Jet.OLEDB.12.0;" & "Data Source =" & MDB_Path & ";" & "Extended Properties=Excel 8.0;"


'判断各表是否存在然后读取相应数据
Set E_RxSchema = E_Connect.OpenSchema(20)

Do Until E_RxSchema.EOF

  If UCase(E_RxSchema("TABLE_TYPE")) = "TABLE" Then
  'Debug.Print E_RxSchema("TABLE_TYPE") & "," & E_RxSchema("TABLE_NAME")
      
      Select Case (E_RxSchema("TABLE_NAME"))
          
        '==================================================================================================================表 Story Drifts
        Case ("Story Drifts")
                  
          '读取时程层间位移角X
          For j = 0 To Num_X - 1
            StrSQL1 = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & name_X(j) & " Max')"
            StrSQL2 = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & name_X(j) & " Min')"
            'Debug.Print StrSQL
            
            Set E_RcdSet1 = New ADODB.Recordset
            E_RcdSet1.Open StrSQL1, E_Connect, 3, 2
            
            Set E_RcdSet2 = New ADODB.Recordset
            E_RcdSet2.Open StrSQL2, E_Connect, 3, 2
            
            If E_RcdSet1.RecordCount <> Num_all Then
              MsgBox "时程" & name_X(j) & "位移角数据不足！"
              Exit Sub
            End If
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
              Temp_a = Abs(E_RcdSet1.Fields("Drift").Value)
              Temp_b = Abs(E_RcdSet2.Fields("Drift").Value)
              
              '取绝对值最大对应的值
              'If Application.WorksheetFunction.Max(Temp_a, Temp_b) > -Application.WorksheetFunction.Min(Temp_a, Temp_b) Then
              '    Sheets("e_E").Cells(i + 2, j * 3 + 10) = Round(1 / Application.WorksheetFunction.Max(Temp_a, Temp_b), 0)
              'Else
              '    Sheets("e_E").Cells(i + 2, j * 3 + 10) = Round(1 / Application.WorksheetFunction.Min(Temp_a, Temp_b), 0)
              'End If
                  
              Sheets("e_E").Cells(i + 2, j * 3 + 10) = Round(1 / Application.WorksheetFunction.Max(Temp_a, Temp_b), 0)
                  
              E_RcdSet1.MoveNext
              E_RcdSet2.MoveNext
            Next i
          Next j
          
          '读取时程层间位移角Y
          For j = 0 To Num_Y - 1
            StrSQL1 = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & name_Y(j) & " Max')"
            StrSQL2 = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & name_Y(j) & " Min')"
            'Debug.Print StrSQL
            
            Set E_RcdSet1 = New ADODB.Recordset
            E_RcdSet1.Open StrSQL1, E_Connect, 3, 2
            
            Set E_RcdSet2 = New ADODB.Recordset
            E_RcdSet2.Open StrSQL2, E_Connect, 3, 2
            
            If E_RcdSet1.RecordCount <> Num_all Then
              MsgBox "时程" & name_Y(j) & "位移角数据不足！"
              Exit Sub
            End If
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
              Temp_a = Abs(E_RcdSet1.Fields("Drift").Value)
              Temp_b = Abs(E_RcdSet2.Fields("Drift").Value)
              
              '取绝对值最大对应的值
              'If Application.WorksheetFunction.Max(Temp_a, Temp_b) > -Application.WorksheetFunction.Min(Temp_a, Temp_b) Then
              '    Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 10) = Round(1 / Application.WorksheetFunction.Max(Temp_a, Temp_b), 0)
              'Else
              '    Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 10) = Round(1 / Application.WorksheetFunction.Min(Temp_a, Temp_b), 0)
              'End If
              
              Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 10) = Round(1 / Application.WorksheetFunction.Max(Temp_a, Temp_b), 0)
                  
                  
              E_RcdSet1.MoveNext
              E_RcdSet2.MoveNext
            Next i
          Next j
                   
       
          
        '==================================================================================================================表 Story Forces
        Case ("Story Forces")
          '读取时程层剪力X
          For j = 0 To Num_X - 1
            StrSQL1 = "Select [Story],[VX],[MY] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & name_X(j) & " Max')"
            StrSQL2 = "Select [Story],[VX],[MY] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & name_X(j) & " Min')"
            'Debug.Print StrSQL
            
            Set E_RcdSet1 = New ADODB.Recordset
            E_RcdSet1.Open StrSQL1, E_Connect, 3, 2
            
            Set E_RcdSet2 = New ADODB.Recordset
            E_RcdSet2.Open StrSQL2, E_Connect, 3, 2
            
            If E_RcdSet1.RecordCount <> Num_all Then
              MsgBox "时程" & name_X(j) & "层剪力数据不足！"
              Exit Sub
            End If
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
              Temp_a = Abs(E_RcdSet1.Fields("VX").Value)
              Temp_b = Abs(E_RcdSet2.Fields("VX").Value)
              Temp_a1 = Abs(E_RcdSet1.Fields("MY").Value)
              Temp_b1 = Abs(E_RcdSet2.Fields("MY").Value)
                            
              Sheets("e_E").Cells(i + 2, j * 3 + 11) = Application.WorksheetFunction.Max(Temp_a, Temp_b)
              Sheets("e_E").Cells(i + 2, j * 3 + 12) = Application.WorksheetFunction.Max(Temp_a1, Temp_b1)
                  
              E_RcdSet1.MoveNext
              E_RcdSet2.MoveNext
            Next i
          Next j
          
          '读取时程层剪力Y
          For j = 0 To Num_Y - 1
            StrSQL1 = "Select [Story],[VY],[MX] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & name_Y(j) & " Max')"
            StrSQL2 = "Select [Story],[VY],[MX] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & name_Y(j) & " Min')"
            'Debug.Print StrSQL
            
            Set E_RcdSet1 = New ADODB.Recordset
            E_RcdSet1.Open StrSQL1, E_Connect, 3, 2
            
            Set E_RcdSet2 = New ADODB.Recordset
            E_RcdSet2.Open StrSQL2, E_Connect, 3, 2
            
            If E_RcdSet1.RecordCount <> Num_all Then
              MsgBox "时程" & name_Y(j) & "层剪力数据不足！"
              Exit Sub
            End If
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
              Temp_a = Abs(E_RcdSet1.Fields("VY").Value)
              Temp_b = Abs(E_RcdSet2.Fields("VY").Value)
              Temp_a1 = Abs(E_RcdSet1.Fields("MX").Value)
              Temp_b1 = Abs(E_RcdSet2.Fields("MX").Value)
              
              '取绝对值最大对应的值
              'If Application.WorksheetFunction.Max(Temp_a, Temp_b) > -Application.WorksheetFunction.Min(Temp_a, Temp_b) Then
              '    Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 10) = Round(1 / Application.WorksheetFunction.Max(Temp_a, Temp_b), 0)
              'Else
              '    Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 10) = Round(1 / Application.WorksheetFunction.Min(Temp_a, Temp_b), 0)
              'End If
              
              Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 11) = Application.WorksheetFunction.Max(Temp_a, Temp_b)
              Sheets("e_E").Cells(i + 2, (j + Num_X) * 3 + 12) = Application.WorksheetFunction.Max(Temp_a1, Temp_b1)
                  
              E_RcdSet1.MoveNext
              E_RcdSet2.MoveNext
            Next i
          Next j
          

        '==================================================================================================================
        Case Else
        
          
        End Select

  End If
  
  E_RxSchema.MoveNext
  
Loop


'关闭Access文件
E_RcdSet1.Close
Set E_RcdSet1 = Nothing
E_RcdSet2.Close
Set E_RcdSet2 = Nothing
E_RxSchema.Close
E_Connect.Close
Set E_Connect = Nothing


'*************************************************************************************************************将平均值最大值标题项写入层分布表
Dim Sum1, Max1, Sum2, Max2, Sum3, Max3

If Num_X <> 0 Then
    '----------------------------平均值X
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y)), Cells(1, 12 + 3 * (Num_X + Num_Y))).MergeCells = True
    Sheets(ela).Cells(1, 10 + 3 * (Num_X + Num_Y)) = "平均值X"
    Sheets(ela).Cells(2, 10 + 3 * (Num_X + Num_Y)) = "层间位移角X"
    Sheets(ela).Cells(2, 11 + 3 * (Num_X + Num_Y)) = "剪力X"
    Sheets(ela).Cells(2, 12 + 3 * (Num_X + Num_Y)) = "倾覆弯矩X"
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y)), Cells(2, 12 + 3 * (Num_X + Num_Y))).Interior.color = Colour
    Temp_Colour = 1 * Temp_Colour
    
    '----------------------------最大值X
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 2)), Cells(1, 12 + 3 * (Num_X + Num_Y + 2))).MergeCells = True
    Sheets(ela).Cells(1, 10 + 3 * (Num_X + Num_Y + 2)) = "最大值X"
    Sheets(ela).Cells(2, 10 + 3 * (Num_X + Num_Y + 2)) = "层间位移角X"
    Sheets(ela).Cells(2, 11 + 3 * (Num_X + Num_Y + 2)) = "剪力X"
    Sheets(ela).Cells(2, 12 + 3 * (Num_X + Num_Y + 2)) = "倾覆弯矩X"
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 2)), Cells(2, 12 + 3 * (Num_X + Num_Y + 2))).Interior.color = Colour
    Temp_Colour = -1 * Temp_Colour
    
    '----------------------------计算
    For i = 1 To Num_all
    
        Sum1 = 0
        Max1 = 1E+64
        Sum2 = 0
        Max2 = 0
        Sum3 = 0
        Max3 = 0
        
        For j = 1 To Num_X
            
            Sum1 = Sum1 + Sheets(ela).Cells(i + 2, 10 + 3 * (j - 1))
            Max1 = Application.WorksheetFunction.Min(Max1, Sheets(ela).Cells(i + 2, 10 + 3 * (j - 1)))
            Sum2 = Sum2 + Sheets(ela).Cells(i + 2, 11 + 3 * (j - 1))
            Max2 = Application.WorksheetFunction.Max(Max2, Sheets(ela).Cells(i + 2, 11 + 3 * (j - 1)))
            Sum3 = Sum3 + Sheets(ela).Cells(i + 2, 12 + 3 * (j - 1))
            Max3 = Application.WorksheetFunction.Max(Max3, Sheets(ela).Cells(i + 2, 12 + 3 * (j - 1)))
        Next j
        
        Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + Num_Y)) = Sum1 / Num_X
        Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + Num_Y + 2)) = Max1
        Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + Num_Y)) = Sum2 / Num_X
        Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + Num_Y + 2)) = Max2
        Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + Num_Y)) = Sum3 / Num_X
        Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + Num_Y + 2)) = Max3
    Next i
Else

End If

If Num_Y <> 0 Then
    '----------------------------平均值Y
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 1)), Cells(1, 12 + 3 * (Num_X + Num_Y + 1))).MergeCells = True
    Sheets(ela).Cells(1, 10 + 3 * (Num_X + Num_Y + 1)) = "平均值Y"
    Sheets(ela).Cells(2, 10 + 3 * (Num_X + Num_Y + 1)) = "层间位移角Y"
    Sheets(ela).Cells(2, 11 + 3 * (Num_X + Num_Y + 1)) = "剪力Y"
    Sheets(ela).Cells(2, 12 + 3 * (Num_X + Num_Y + 1)) = "倾覆弯矩Y"
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 1)), Cells(2, 12 + 3 * (Num_X + Num_Y + 1))).Interior.color = Colour
    Temp_Colour = 1 * Temp_Colour
    
    '----------------------------最大值Y
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 3)), Cells(1, 12 + 3 * (Num_X + Num_Y + 3))).MergeCells = True
    Sheets(ela).Cells(1, 10 + 3 * (Num_X + Num_Y + 3)) = "最大值Y"
    Sheets(ela).Cells(2, 10 + 3 * (Num_X + Num_Y + 3)) = "层间位移角Y"
    Sheets(ela).Cells(2, 11 + 3 * (Num_X + Num_Y + 3)) = "剪力Y"
    Sheets(ela).Cells(2, 12 + 3 * (Num_X + Num_Y + 3)) = "倾覆弯矩Y"
    '加背景色
    If Temp_Colour > 0 Then
      Colour = 10091441
    Else
      Colour = 6750207
    End If
    Sheets(ela).range(Cells(1, 10 + 3 * (Num_X + Num_Y + 3)), Cells(2, 12 + 3 * (Num_X + Num_Y + 3))).Interior.color = Colour
    Temp_Colour = -1 * Temp_Colour
    
    '----------------------------计算
    For i = 1 To Num_all
    
        Sum1 = 0
        Max1 = 1E+64
        Sum2 = 0
        Max2 = 0
        Sum3 = 0
        Max3 = 0
        
        For j = 1 To Num_Y
            
            Sum1 = Sum1 + Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + j - 1))
            Max1 = Application.WorksheetFunction.Min(Max1, Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + j - 1)))
            Sum2 = Sum2 + Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + j - 1))
            Max2 = Application.WorksheetFunction.Max(Max2, Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + j - 1)))
            Sum3 = Sum3 + Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + j - 1))
            Max3 = Application.WorksheetFunction.Max(Max3, Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + j - 1)))
        Next j
        
        Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + Num_Y + 1)) = Sum1 / Num_Y
        Sheets(ela).Cells(i + 2, 10 + 3 * (Num_X + Num_Y + 3)) = Max1
        Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + Num_Y + 1)) = Sum2 / Num_Y
        Sheets(ela).Cells(i + 2, 11 + 3 * (Num_X + Num_Y + 3)) = Max2
        Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + Num_Y + 1)) = Sum3 / Num_Y
        Sheets(ela).Cells(i + 2, 12 + 3 * (Num_X + Num_Y + 3)) = Max3
        
    Next i
Else

End If



'*************************************************************************************************************写入反应谱数据

Sheets(ela).Cells(Num_X + Num_Y + 6, 1) = "平均值"
Sheets(ela).Cells(Num_X + Num_Y + 7, 1) = "最大值"
Sheets(ela).Cells(Num_X + Num_Y + 8, 1) = "反应谱"
Sheets(ela).Cells(Num_X + Num_Y + 9, 1) = "65%反应谱"
Sheets(ela).Cells(Num_X + Num_Y + 10, 1) = "135%反应谱"
Sheets(ela).Cells(Num_X + Num_Y + 11, 1) = "80%反应谱"
Sheets(ela).Cells(Num_X + Num_Y + 12, 1) = "120%反应谱"

'=============================================================================================================================将反应谱数据转入
Sheets(ela).Cells(1, 10 + (Num_X + Num_Y + 4) * 3) = "反应谱"
Sheets(ela).range(Cells(1, (Num_X + Num_Y + 4) * 3 + 10), Cells(1, (Num_X + Num_Y + 4) * 3 + 23)).MergeCells = True
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 10) = "X层间位移角"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 13) = "Y层间位移角"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 11) = "X剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 14) = "Y剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 12) = "X倾覆弯矩"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 15) = "Y倾覆弯矩"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 16) = "65%X剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 17) = "135%X剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 18) = "80%X剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 19) = "120%X剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 20) = "65%Y剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 21) = "135%Y剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 22) = "80%Y剪力"
Sheets(ela).Cells(2, (Num_X + Num_Y + 4) * 3 + 23) = "120%Y剪力"

'加背景色
If Temp_Colour > 0 Then
  Colour = 10091441
Else
  Colour = 6750207
End If

Sheets(ela).range(Cells(1, (Num_X + Num_Y + 4) * 3 + 10), Cells(2, (Num_X + Num_Y + 4) * 3 + 23)).Interior.color = Colour
Temp_Colour = -1 * Temp_Colour


'层间位移角
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 10), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 10)).Value = Sheets("d_E").range("Z3:" & "Z" & Num_all + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 13), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 13)).Value = Sheets("d_E").range("AD3:" & "AD" & Num_all + 2).Value
'剪力
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 11), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 11)).Value = Sheets("d_E").range("J3:" & "J" & Num_all + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 14), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 14)).Value = Sheets("d_E").range("N3:" & "N" & Num_all + 2).Value
'弯矩
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 12), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 12)).Value = Sheets("d_E").range("K3:" & "K" & Num_all + 2).Value
Sheets(ela).range(Sheets(ela).Cells(3, (Num_X + Num_Y + 4) * 3 + 15), Sheets(ela).Cells(Num_all + 2, (Num_X + Num_Y + 4) * 3 + 15)).Value = Sheets("d_E").range("O3:" & "O" & Num_all + 2).Value

For i = 1 To Num_all
'X正负35%剪力
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 16) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 11) * 0.65
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 17) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 11) * 1.35
'X正负20%剪力
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 18) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 11) * 0.8
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 19) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 11) * 1.2
'Y正负35%剪力
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 20) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 14) * 0.65
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 21) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 14) * 1.35
'Y正负20%剪力
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 22) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 14) * 0.8
Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 23) = Sheets(ela).Cells(i + 2, (Num_X + Num_Y + 4) * 3 + 14) * 1.2
Next i

'=============================================================================================================================填写汇总表格

'读取反应谱的基底剪力
Sheets(ela).Cells(Num_X + Num_Y + 8, 2) = Sheets("d_E").Cells(3, 10)
Sheets(ela).Cells(Num_X + Num_Y + 8, 5) = Sheets("d_E").Cells(3, 14)
Sheets(ela).Cells(Num_X + Num_Y + 9, 2) = Sheets("d_E").Cells(3, 10) * 0.65
Sheets(ela).Cells(Num_X + Num_Y + 9, 5) = Sheets("d_E").Cells(3, 14) * 0.65
Sheets(ela).Cells(Num_X + Num_Y + 10, 2) = Sheets("d_E").Cells(3, 10) * 1.35
Sheets(ela).Cells(Num_X + Num_Y + 10, 5) = Sheets("d_E").Cells(3, 14) * 1.35
Sheets(ela).Cells(Num_X + Num_Y + 11, 2) = Sheets("d_E").Cells(3, 10) * 0.8
Sheets(ela).Cells(Num_X + Num_Y + 11, 5) = Sheets("d_E").Cells(3, 14) * 0.8
Sheets(ela).Cells(Num_X + Num_Y + 12, 2) = Sheets("d_E").Cells(3, 10) * 1.2
Sheets(ela).Cells(Num_X + Num_Y + 12, 5) = Sheets("d_E").Cells(3, 14) * 1.2

'读取各时程下基底剪力，汇总至表格
'X向
For i = 1 To Num_X

    '基底剪力汇总
    Sheets(ela).Cells(5 + i, 2).Value = Sheets(ela).Cells(3, 11 + (i - 1) * 3)
    
    '时程结果与反应谱结果的比值
    If Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 2) = "" Then
        'Debug.Print "缺少反应谱数据！"
    Else
        Sheets(ela).Cells(5 + i, 3).Value = Round(Sheets(ela).Cells(5 + i, 2).Value / Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 2).Value, 3)
        
    End If
    
    '位移角汇总
    '最大位移角所在行数
    Dim RRX, RRY As Integer
    'RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (i - 1) * 3)))(1)
    '使用excel自带函数求最值所在位置，避免with报错
    RRX = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (i - 1) * 3))), Sheets(ela).range(Sheets(ela).Cells(3, 10 + (i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (i - 1) * 3)), 0)
    'Debug.Print RRX
    '将最大位移角及构件编号写入表格
    Sheets(ela).Cells(5 + i, 4) = Worksheets(ela).Cells(RRX + 2, 10 + (i - 1) * 3)
    Worksheets(ela).Cells(RRX + 2, 10 + (i - 1) * 3).Interior.ColorIndex = 4
    
Next i

'Y向
For i = 1 To Num_Y

    '基底剪力汇总
    Sheets(ela).Cells(5 + Num_X + i, 5).Value = Sheets(ela).Cells(3, 11 + (Num_X + i - 1) * 3)
    
    '时程结果与反应谱结果的比值
    If Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 5) = "" Then
        'Debug.Print "缺少反应谱数据！"
    Else
        Sheets(ela).Cells(5 + Num_X + i, 6).Value = Round(Sheets(ela).Cells(5 + Num_X + i, 5).Value / Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 5).Value, 3)
        
    End If
    
    '位移角汇总
    '最大位移角所在行数
    'RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + i - 1) * 3)))(1)
    RRX = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + i - 1) * 3))), Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + i - 1) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + i - 1) * 3)), 0)
    '将最大位移角及构件编号写入表格
    Sheets(ela).Cells(Num_X + 5 + i, 7) = Worksheets(ela).Cells(RRX + 2, 10 + (Num_X + i - 1) * 3)
    Worksheets(ela).Cells(RRX + 2, 10 + (Num_X + i - 1) * 3).Interior.ColorIndex = 4
    
Next i

'平均值，最大值
For i = 1 To 2

    '基底剪力汇总
    Sheets(ela).Cells(5 + Num_X + Num_Y + i, 2).Value = Sheets(ela).Cells(3, 11 + (Num_X + Num_Y + 2 * (i - 1)) * 3)
    Sheets(ela).Cells(5 + Num_X + Num_Y + i, 5).Value = Sheets(ela).Cells(3, 14 + (Num_X + Num_Y + 2 * (i - 1)) * 3)
    
    '时程结果与反应谱结果的比值
    If Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 2) = "" Or Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 5) = "" Then
        'Debug.Print "缺少反应谱数据！"
    Else
        Sheets(ela).Cells(5 + Num_X + Num_Y + i, 3).Value = Round(Sheets(ela).Cells(5 + Num_X + Num_Y + i, 2).Value / Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 2).Value, 3)
        Sheets(ela).Cells(5 + Num_X + Num_Y + i, 6).Value = Round(Sheets(ela).Cells(5 + Num_X + Num_Y + i, 5).Value / Sheets(ela).Cells(Num_X + Num_Y + 2 + 6, 5).Value, 3)
        
    End If
    
    '位移角汇总,X
    If Num_X <> 0 Then
        '最大位移角所在行数
        'RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3)))(1)
        RRX = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3))), Sheets(ela).range(Sheets(ela).Cells(3, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3)), 0)
        '将最大位移角及构件编号写入表格
        Sheets(ela).Cells(Num_X + Num_Y + 5 + i, 4) = Worksheets(ela).Cells(RRX + 2, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3)
        Worksheets(ela).Cells(RRX + 2, 10 + (Num_X + Num_Y + 2 * (i - 1)) * 3).Interior.ColorIndex = 4
    End If
    
    
    '位移角汇总，Y
    If Num_Y <> 0 Then
        '最大位移角所在行数
        'RRX = IndexMinofRange(Sheets(ela).range(Sheets(ela).Cells(3, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3)))(1)
        RRX = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Sheets(ela).range(Sheets(ela).Cells(3, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3))), Sheets(ela).range(Sheets(ela).Cells(3, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3), Sheets(ela).Cells(Num_all + 2, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3)), 0)
        '将最大位移角及构件编号写入表格
        Sheets(ela).Cells(Num_X + Num_Y + 5 + i, 7) = Worksheets(ela).Cells(RRX + 2, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3)
        Worksheets(ela).Cells(RRX + 2, 13 + (Num_X + Num_Y + 2 * (i - 1)) * 3).Interior.ColorIndex = 4
    End If
    
Next i

'加背景色
Call AddShadow(ela, "A2:A" & Num_X + Num_Y + 12, 10092441)
Call AddShadow(ela, "B4:G5,C2:C3,E2:E3", 10092441)
Call AddShadow(ela, "B2:B3,D2:D3,F2:F3", 6750207)
Call AddShadow(ela, "B6:G" & Num_X + Num_Y + 12, 6750207)

'所有单元格宽度自动调整
Sheets(ela).Cells.EntireColumn.AutoFit

End Sub

Function IndexMinofRange(index_Range As range)
Dim Min, R, C As Integer
Min = WorksheetFunction.Min(index_Range)
R = index_Range.Find(Min, LookIn:=xlValues).Row
'C = index_Range.Find(Min, After:=index_Range.Cells(index_Range.Rows.Count, index_Range.Columns.Count), LookIn:=xlValues).Columns
IndexMinofRange = Array(Min, R) ', C)
End Function


