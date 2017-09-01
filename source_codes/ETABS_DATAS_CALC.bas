Attribute VB_Name = "ETABS_DATAS_CALC"

'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                           ETABS数据处理                              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2015/04/19
'更新内容:
'1.添加剪重比修正位移角


'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2015/04/15
'更新内容:
'1.修改Num_All

'////////////////////////////////////////////////////////////////////////////////////////////

'更新时间:2014/06/11
'更新内容:
'1.添加range是否为空判断,解决缺少数据时with报错问题


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/4
'更新内容:
'1.增加general表最小刚度比计算
'1.增加总质量计算


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/11
'更新内容:
'1.补充楼层质量数据


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/8
'更新内容:
'1.数据处理代码，根据调用函数时传递的变量（读取MDB表格名）来进行相应数据处理

Public Sub ETABS_DATA_CALC(T_Name As String)

Select Case (T_Name)

    '======================================================================================表 Modal Direction Factors
    Case "Modal Direction Factors"
      '计算周期比
      Sheets("g_E").Cells(38, 4).FormulaArray = "=INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37>0.5,))/INDEX($D$28:$D$37,MATCH(TRUE,$G$28:$G$37<0.5,))"
      Sheets("g_E").Cells(38, 5).Formula = "=if(d38<0.85,""< 0.85"",""> 0.85"")"
      
      
      
    '======================================================================================表 Story Drifts
    Case "Story Drifts"
      '----------------------------------------------------读取最大层间位移角、所在楼层及工况
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  风荷载
      Sheets("g_E").Cells(10, 5).Formula = "=1&"" / ""&MIN(d_E!AC:AC)"
      Sheets("g_E").Cells(10, 7).Formula = "=1&"" / ""&MIN(d_E!AG:AG)"
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震
      Sheets("g_E").Cells(11, 5).Formula = "=1&"" / ""&MIN(d_E!Z:Z)"
      Sheets("g_E").Cells(11, 7).Formula = "=1&"" / ""&MIN(d_E!AD:AD)"
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震+
      Sheets("g_E").Cells(12, 5).Formula = "=1&"" / ""&MIN(d_E!AA:AA)"
      Sheets("g_E").Cells(12, 7).Formula = "=1&"" / ""&MIN(d_E!AE:AE)"
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  地震-
      Sheets("g_E").Cells(13, 5).Formula = "=1&"" / ""&MIN(d_E!AB:AB)"
      Sheets("g_E").Cells(13, 7).Formula = "=1&"" / ""&MIN(d_E!AF:AF)"
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  最大层间位移角
      Dim Temp_min As Integer
      Temp_min = Application.WorksheetFunction.Min(range("d_E!Z3:Z" & Num_all + 2 & ",d_E!AC3:AD" & Num_all + 2 & " ,d_E!AG3:AG" & Num_all + 2))
      Sheets("g_E").Cells(14, 4).Formula = " 1 / " & Temp_min & ""
      '=1&"/" & 5
      
      '定义最大层间位移角查询区域
      Set iRng_X = Application.Union(range("d_E!Z:Z"), range("d_E!AC:AD"), range("d_E!AG:AG"))
      
      '定义查询变量
      Dim i_Min As Double, i_Row As Integer, i_Col As Integer
      Dim i_Temp As range
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  XY向
      '查询区域内最大层间位移角（实际是查询层间位移角的最小分母）
      i_Min = WorksheetFunction.Min(iRng_X)
      
      Set i_Temp = iRng_X.Find(i_Min, After:=iRng_X.Cells(iRng_X.Rows.Count, iRng_X.Columns.Count), LookIn:=xlValues, lookat:=xlWhole)

      If Not i_Temp Is Nothing Then
        '返回最大层间位移角所在行号、列号
        i_Row = i_Temp.Row
        i_Col = i_Temp.column
        '返回最大层间位移角所在层，及其工况
        Sheets("g_E").Cells(15, 7) = Sheets("d_E").Cells(i_Row, 1)
        Sheets("g_E").Cells(15, 5) = Sheets("d_E").Cells(2, i_Col)
      End If
      
      '返回最大层间位移角所在行号、列号
'      On Error Resume Next
'      i_Row = iRng_X.Find(What:=i_Min, lookat:=xlWhole).Row
'      i_Col = iRng_X.Find(What:=i_Min, lookat:=xlWhole).column
      '返回最大层间位移角所在层，及其工况
'      Sheets("g_E").Cells(15, 7) = Sheets("d_E").Cells(i_Row, 1)
'      Sheets("g_E").Cells(15, 5) = Sheets("d_E").Cells(2, i_Col)
      
      
    '=======================================================================================表 Story Stiffness
    Case "Story Stiffness"
      '不读取输出刚度和刚度比，自行计算
        For i = 1 To Num_all
        Sheets("d_E").Cells(i + 2, 4).Formula = "=RC[6]*RC[22]"
        Sheets("d_E").Cells(i + 2, 5).Formula = "=RC[9]*RC[25]"
        Sheets("d_E").Cells(i + 2, 2).Formula = "=RC[2]/R[1]C[2]"
        Sheets("d_E").Cells(i + 2, 3).Formula = "=RC[2]/R[1]C[2]"
        Next i
        Sheets("d_E").Cells(Num_all + 2, 2) = 1
        Sheets("d_E").Cells(Num_all + 2, 3) = 1
      '------------------------------------------------------读取最小刚度比
      Sheets("g_E").Cells(22, 5).Formula = "=MIN(d_E!B3:B" & Num_all + 1 & ")"
      Sheets("g_E").Cells(22, 7).Formula = "=MIN(d_E!C3:C" & Num_all + 1 & ")"
      
      Debug.Print ""
      
    
    '=======================================================================================表 Story Max/Avg Displacements
    Case "Story Max/Avg Displacements"
      '------------------------------------------------------读取最大位移比
      Sheets("g_E").Cells(16, 4).Formula = "=MAX(d_E!AH3:AM" & Num_all + 2 & ")"
      
      Sheets("g_E").Cells(17, 7).FormulaArray = "=INDIRECT(""d_E!R"" & MIN(IF(d_E!AH3:d_E!AM" & Num_all + 2 & "=D16,ROW(d_E!AH3:d_E!AM" & Num_all + 2 & "))) & ""C1"",False)"
      Sheets("g_E").Cells(17, 5).FormulaArray = "=INDIRECT(""d_E!R2C"" & MIN(IF(d_E!AH3:d_E!AM" & Num_all + 2 & "=D16,COLUMN(d_E!AH3:d_E!AM" & Num_all + 2 & "))),FALSE)"
      
      
    '=======================================================================================表 Diaphragm CM Displacements
    Case "Diaphragm CM Displacements"
      '------------------------------------------------------读取最大位移比
      
      For i = 1 To Num_all
         Sheets("d_E").Cells(i + 2, 35).Formula = "=RC[35]/RC[-16]"
         Sheets("d_E").Cells(i + 2, 36).Formula = "=RC[35]/RC[-16]"
         Sheets("d_E").Cells(i + 2, 38).Formula = "=RC[34]/RC[-15]"
         Sheets("d_E").Cells(i + 2, 39).Formula = "=RC[34]/RC[-15]"
      Next i
      
      Sheets("g_E").Cells(16, 4).Formula = "=MAX(d_E!AH3:AM" & Num_all + 2 & ")"
      Sheets("g_E").Cells(17, 7).FormulaArray = "=INDIRECT(""d_E!R"" & MIN(IF(d_E!AH3:d_E!AM" & Num_all + 2 & "=D16,ROW(d_E!AH3:d_E!AM" & Num_all + 2 & "))) & ""C1"",False)"
      Sheets("g_E").Cells(17, 5).FormulaArray = "=INDIRECT(""d_E!R2C"" & MIN(IF(d_E!AH3:d_E!AM" & Num_all + 2 & "=D16,COLUMN(d_E!AH3:d_E!AM" & Num_all + 2 & "))),FALSE)"
      
    'Diaphragm CM Displacements
      
    '=======================================================================================表 Story Forces
    Case "Story Forces"
      '风X向剪力
      Sheets("g_E").Cells(42, 4).Formula = "=d_E!F3"
      '风X向弯矩
      Sheets("g_E").Cells(42, 6).Formula = "=d_E!G3"
      '风Y向剪力
      Sheets("g_E").Cells(43, 4).Formula = "=d_E!H3"
      '风Y向弯矩
      Sheets("g_E").Cells(43, 6).Formula = "=d_E!I3"
      '地震X向剪力
      Sheets("g_E").Cells(44, 4).Formula = "=d_E!J3"
      '地震X向弯矩
      Sheets("g_E").Cells(44, 6).Formula = "=d_E!K3"
      '地震Y向剪力
      Sheets("g_E").Cells(45, 4).Formula = "=d_E!N3"
      '地震Y向弯矩
      Sheets("g_E").Cells(45, 6).Formula = "=d_E!O3"
    
    
    '======================================================================================表 Shear Gravity Ratios
    Case "Shear Gravity Ratios"
      '读取最小剪重比
      Sheets("g_E").Cells(24, 5).Formula = "=MIN(d_E!L3:L" & Num_all + 2 & ")"
      Sheets("g_E").Cells(25, 5).Formula = "=MIN(d_E!P3:P" & Num_all + 2 & ")"
    
    
    '======================================================================================表 Mass Summary by Story
    Case ("Mass Summary by Story")
      
      Sheets("d_E").Cells(3, 55) = 1
      
      For i = 2 To Num_all
          If Sheets("d_E").Cells(i + 1, 54) <> 0 Then
          Sheets("d_E").Cells(i + 2, 55).Formula = "=RC[-1]/R[-1]C[-1]"
          End If
      Next i
      
      Sheets("g_E").Cells(7, 7).Formula = "=SUM(d_E!BB3:BB" & Num_all + 2 & ")"

    '======================================================================================ETABS位移角剪重比修正
    Case ("ETABSMOB")
      'Debug.Print OUTReader_Main.RatioLimitY_TextBox.Value / Sheets("d_E").Cells(3, 16).Value
      If Round(Sheets("d_E").Cells(3, 12).Value, 2) < Round(OUTReader_Main.RatioLimitX_TextBox.Value, 2) Then
        For i = 1 To 53
            Sheets("d_E").Cells(i + 2, 13).Value = OUTReader_Main.RatioLimitX_TextBox.Value
            a = Sheets("d_E").Cells(i + 2, 26).Value
            'Debug.Print a
            Sheets("d_E").Cells(i + 2, 26) = Round(a * Sheets("d_E").Cells(3, 12).Value / OUTReader_Main.RatioLimitX_TextBox.Value, 0)
        Next i
    
      End If
      If Round(Sheets("d_E").Cells(3, 16).Value, 2) < Round(OUTReader_Main.RatioLimitY_TextBox.Value, 2) Then
        'Debug.Print "ssss"
        For i = 1 To 53
            Sheets("d_E").Cells(i + 2, 17).Value = OUTReader_Main.RatioLimitY_TextBox.Value
            b = Sheets("d_E").Cells(i + 2, 30).Value
            'Debug.Print a
            Sheets("d_E").Cells(i + 2, 30) = Round(b * Sheets("d_E").Cells(3, 16).Value / OUTReader_Main.RatioLimitY_TextBox.Value, 0)
        Next i
    
      End If
      MsgBox "请注意不要重复修正！"
    Case Else

      
End Select


End Sub

