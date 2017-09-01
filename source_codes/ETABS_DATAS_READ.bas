Attribute VB_Name = "ETABS_DATAS_READ"
'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                           ETABS数据读取                              ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2015/4/15
'更新内容:
'1.将周期因子改为质量参与系数;


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/4/4
'更新内容:
'1.增加调整后剪力提取 表 Frame Shear Ratios In Dual Systems And Modifiers 和倾覆弯矩提取 表 Frame Overturning Moments In Dual Systems
'2.增加general表刚度比提取调用
'3.增加版本和时间

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/29
'更新内容:
'1.添加针对V9的代码:位移和框架剪力调整；
'2.对exit sub 进行了调整，直接exit太粗暴了~


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/28
'更新内容:
'1.添加针对V9的代码;
'2.添加了刚度比求解（放在了刚度比部分）


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/11
'更新内容:
'1.某一工况下位移比数据不区分X向Y向，选择条件语句中去除方向条件 [Direction] IN ('X')或[Direction] IN ('Y')
'2.etabs位移比输出时对未定义隔板的楼层不输出数据，全为0，位移比为空，请注意
'3.补充楼层质量数据


'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/6
'更新内容:
'1.ETABS数据提取，目前写好提取Story Data, Modal Direction Factors, Modal Participating Mass Ratios,
'Story Drifts, Story Max/Avg Displacements, Story Forces, Story Stiffness, Shear Gravity Ratios.
'2.目前缺少刚度比，质量分布等信息，没有在etabs2013输出选项中找到


Public Sub ETABS_DATA_READ(MDB_Path As String)

'定义变量
Dim E_Connect As ADODB.Connection
Dim E_RcdSet As ADODB.Recordset
Dim E_RxSchema As ADODB.Recordset

Dim StrSQL As String

'判断Access文件是否存在
If Dir(MDB_Path) = " " Then
  MsgBox "MDB文件不存在！请核实！", vbExclamation, "无法连接数据库"
  Exit Sub
End If


'使用ADO连接Access文件
'对于Access 2007 及高版本EXCEL
Set E_Connect = New ADODB.Connection
E_Connect.CursorLocation = adUseClient
E_Connect.Open ConnectionString:="Provider=Microsoft.Ace.OLEDB.12.0;" & "Data Source =" & MDB_Path & ";" '& "Extended Properties=Excel 12.0;"
'对于早期版本的Access和Excel使用
'myConnect.Open ConnectionString:="Provider=Microsoft.Jet.OLEDB.12.0;" & "Data Source =" & MDB_Path & ";" & "Extended Properties=Excel 8.0;"


'================================================================================================================表 Story Data
  
  Dim Story_N() As String
  
  '读取楼层名
    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
        StrSQL = "Select [Name],[Height] From [Story Data] Order By [Elevation] ASC"
    ElseIf OUTReader_Main.Option_E_V9 Then
        StrSQL = "Select [Story],[Height] From [Story Data] Order By [Elevation] ASC"
        'Debug.Print StrSQL
    End If
        'StrSQL = "Select [Name],[Height] From [Story Data] Order By [Elevation] ASC"
  
  Set E_RcdSet = New ADODB.Recordset
  E_RcdSet.Open StrSQL, E_Connect, 3, 2
  
  Num_all = E_RcdSet.RecordCount - 1
  ReDim Story_N(Num_all, 2)
  For i = 0 To Num_all
    Story_N(i, 0) = i
    
    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
        Story_N(i, 1) = E_RcdSet.Fields("Name").Value
    ElseIf OUTReader_Main.Option_E_V9 Then
        Story_N(i, 1) = E_RcdSet.Fields("Story").Value
        'Debug.Print StrSQL
    End If
    
    'Story_N(i, 1) = E_RcdSet.Fields("Name").Value
    Story_N(i, 2) = E_RcdSet.Fields("Height").Value
    'Debug.Print Story_N(i, 0) & "," & Story_N(i, 1) & "," & Story_N(i, 2)
    E_RcdSet.MoveNext
  Next i
  
  '写入楼层编号和层高
  For i = 1 To Num_all
    
    Sheets("d_E").Cells(i + 2, 1) = Story_N(i, 0)
    Sheets("d_E").Cells(i + 2, 60) = Story_N(i, 2)
    
  Next i
  
  If OUTReader_Main.ComboBox_SPEC_X.Text = "" And OUTReader_Main.ComboBox_SPEC_XEcc.Text = "" And OUTReader_Main.ComboBox_SPEC_Y.Text = "" And OUTReader_Main.ComboBox_Wind_X.Text = "" And OUTReader_Main.ComboBox_Wind_Y.Text = "" Then
    Exit Sub
  End If
            
            
'================================================================================================================表 Program Control

If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别

    StrSQL = "Select [ProgramName],[Version],[Level] From [Program Control]"
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    Sheets("g_E").Cells(4, 4) = E_RcdSet.Fields("ProgramName").Value & " V" & E_RcdSet.Fields("Version").Value & " " & E_RcdSet.Fields("Level").Value
    
ElseIf OUTReader_Main.Option_E_V9 Then

    Sheets("g_E").Cells(4, 4) = "ETABS V9"
    'Debug.Print StrSQL
End If

Sheets("g_E").Cells(4, 7).Formula = "=Today()"
        

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%数据结果读取
'判断各表是否存在然后读取相应数据
Set E_RxSchema = E_Connect.OpenSchema(20)

Do Until E_RxSchema.EOF

  If UCase(E_RxSchema("TABLE_TYPE")) = "TABLE" Then
  'Debug.Print E_RxSchema("TABLE_TYPE") & "," & E_RxSchema("TABLE_NAME")
      
      Select Case (E_RxSchema("TABLE_NAME"))
      
      
        '===============================================================================================================表 Modal Direction Factors
        Case ("Modal Direction Factors")
          Dim N_M As Long
          
          '读取振型及方向因子
          StrSQL = "Select [Period],[UX],[UY],[RZ] From [Modal Direction Factors]" ' Order By [Mode] ASC"
          
          Set E_RcdSet = New ADODB.Recordset
          E_RcdSet.Open StrSQL, E_Connect, 3, 2
          
          'N_M = E_RcdSet.RecordCount
          N_M = 10
          Sheets("g_E").Cells(38, 7) = N_M
          
          If N_M > 10 Then N_M = 10
          
          For i = 1 To N_M
            Sheets("g_E").Cells(i + 27, 4) = E_RcdSet.Fields("Period").Value
            Sheets("g_E").Cells(i + 27, 6) = Format(E_RcdSet.Fields("UX").Value, "0.00") & "+" & Format(E_RcdSet.Fields("UY").Value, "0.00")
            Sheets("g_E").Cells(i + 27, 7) = E_RcdSet.Fields("RZ").Value
            E_RcdSet.MoveNext
          Next i
          
          Call ETABS_DATA_CALC("Modal Direction Factors")
          
        
        '===============================================================================================================表 Modal Participating Mass Ratios
        Case ("Modal Participating Mass Ratios")
          '读取振型质量参与系数
          StrSQL = "Select [Mode],[UX],[UY],[RZ] ,[SumUX],[SumUY] From [Modal Participating Mass Ratios]"
          
          Set E_RcdSet = New ADODB.Recordset
          E_RcdSet.Open StrSQL, E_Connect, 3, 2
          For i = 1 To 10 '----------------------------------------------------------------------------------------------------------------------------------------------------------添加
            Sheets("g_E").Cells(i + 27, 6) = Format(E_RcdSet.Fields("UX").Value, "0.00") & "+" & Format(E_RcdSet.Fields("UY").Value, "0.00")
            Sheets("g_E").Cells(i + 27, 7) = Format(E_RcdSet.Fields("RZ").Value, "0.00")
            E_RcdSet.MoveNext
          Next i
          
          E_RcdSet.MoveLast
          
          Sheets("g_E").Cells(39, 5) = E_RcdSet.Fields("SumUX")
          Sheets("g_E").Cells(39, 7) = E_RcdSet.Fields("SumUY")
          
          Sheets("g_E").Cells(27, 6) = "质量参与系数（X+Y）"
          Sheets("g_E").Cells(27, 7) = "质量参与系数（Z）"
          

          
                
          
        '==================================================================================================================表 Story Drifts
        Case ("Story Drifts")
                  
          '读取水平地震力层间位移角X
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftX] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
                'Debug.Print StrSQL
            End If
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX位移角数据不足！~"
            Else
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                  
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 26) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 26) = Round(1 / E_RcdSet.Fields("DriftX").Value, 0)
                        'Debug.Print StrSQL
                    End If
                    
                  'Sheets("d_E").Cells(i + 2, 26) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            End If
          End If
          
          '读取水平地震力层间位移角X+
          If OUTReader_Main.ComboBox_SPEC_XEcc.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftX] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & "')"
            'Debug.Print StrSQL
            End If
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX+位移角数据不足！"
            Else
            
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 27) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 27) = Round(1 / E_RcdSet.Fields("DriftX").Value, 0)
                        'Debug.Print StrSQL
                    End If
                    
                  'Sheets("d_E").Cells(i + 2, 27) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            End If
                    
            '读取水平地震力层间位移角X-
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftX] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & "')"
            'Debug.Print StrSQL
            End If
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX-位移角数据不足！"
            Else

            
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 28) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 28) = Round(1 / E_RcdSet.Fields("DriftX").Value, 0)
                        'Debug.Print StrSQL
                    End If
                  'Sheets("d_E").Cells(i + 2, 28) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            
            End If
          End If
          
          '读取风荷载层间位移角X
          If OUTReader_Main.ComboBox_Wind_X.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftX] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [Load] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            'Debug.Print StrSQL
            End If
            
               'StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift X') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "WX位移角数据不足！~"
            Else
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                  
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 29) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 29) = Round(1 / E_RcdSet.Fields("DriftX").Value, 0)
                        'Debug.Print StrSQL
                    End If
                    
                  'Sheets("d_E").Cells(i + 2, 29) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            End If
          End If
          
          '读取水平地震力层间位移角Y
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftY] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
            'Debug.Print StrSQL
            End If
            
                'StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY位移角数据不足！"
            Else
            
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                  
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 30) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 30) = Round(1 / E_RcdSet.Fields("DriftY").Value, 0)
                        'Debug.Print StrSQL
                    End If
                    
                  'Sheets("d_E").Cells(i + 2, 30) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            End If
          End If
          
          '读取水平地震力层间位移角Y+
          If OUTReader_Main.ComboBox_SPEC_YEcc.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftY] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & "')"
            'Debug.Print StrSQL
            End If
            
                'StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY+位移角数据不足！"
            Else

            
                For i = Num_all To 1 Step -1
                  'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
                
                    If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                        Sheets("d_E").Cells(i + 2, 31) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                    ElseIf OUTReader_Main.Option_E_V9 Then
                        Sheets("d_E").Cells(i + 2, 31) = Round(1 / E_RcdSet.Fields("DriftY").Value, 0)
                        'Debug.Print StrSQL
                    End If
                    
                  'Sheets("d_E").Cells(i + 2, 31) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                  E_RcdSet.MoveNext
                Next i
            
            End If
            
            '读取水平地震力层间位移角Y-
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftY] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & "')"
            'Debug.Print StrSQL
            End If
            
                'StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY-位移角数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")
              
                If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                    Sheets("d_E").Cells(i + 2, 32) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                ElseIf OUTReader_Main.Option_E_V9 Then
                    Sheets("d_E").Cells(i + 2, 32) = Round(1 / E_RcdSet.Fields("DriftY").Value, 0)
                    'Debug.Print StrSQL
                End If
                
              'Sheets("d_E").Cells(i + 2, 32) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          '读取风荷载层间位移角Y
          If OUTReader_Main.ComboBox_Wind_Y.Text <> "" Then
            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            ElseIf OUTReader_Main.Option_E_V9 Then
                StrSQL = "Select [Story],[DriftY] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [Load] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            'Debug.Print StrSQL
            End If
            
                'StrSQL = "Select [Story],[Drift] From [Story Drifts] Where [Item] IN ('Max Drift Y') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "WY位移角数据不足！~"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Drift")

                If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
                    Sheets("d_E").Cells(i + 2, 33) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
                ElseIf OUTReader_Main.Option_E_V9 Then
                    Sheets("d_E").Cells(i + 2, 33) = Round(1 / E_RcdSet.Fields("DriftY").Value, 0)
                    'Debug.Print StrSQL
                End If
                
              'Sheets("d_E").Cells(i + 2, 33) = Round(1 / E_RcdSet.Fields("Drift").Value, 0)
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          Call ETABS_DATA_CALC("Story Drifts")
          
        '==================================================================================================================表 Story Max/Avg Displacements '----------------------------------V9不能输出
        Case ("Story Max/Avg Displacements")
          '读取规定水平地震力位移X
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 18) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 34) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取规定水平地震力位移X+
          If OUTReader_Main.ComboBox_SPEC_XEcc.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX+位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 19) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 35) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
                      
            '读取规定水平地震力位移X-
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX-位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 20) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 36) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取风荷载位移X
          If OUTReader_Main.ComboBox_Wind_X.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WX位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 21) = E_RcdSet.Fields("Average").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取规定水平地震力位移Y
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EY位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 22) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 37) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取规定水平地震力位移Y+
          If OUTReader_Main.ComboBox_SPEC_YEcc.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where  [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EY+位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 23) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 38) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
                    
            '读取规定水平地震力位移Y-
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EY-位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 24) = E_RcdSet.Fields("Average").Value
              Sheets("d_E").Cells(i + 2, 39) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          '读取风荷载位移Y
          If OUTReader_Main.ComboBox_Wind_Y.Text <> "" Then
            StrSQL = "Select [Story],[Average],[Ratio] From [Story Max/Avg Displacements] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WY位移数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 25) = E_RcdSet.Fields("Average").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          Call ETABS_DATA_CALC("Story Max/Avg Displacements")
          
        '==================================================================================================================表 Story Forces
        Case ("Story Forces")
          '读取水平地震剪力X、"弯矩Y"
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
            StrSQL = "Select [Story],[VX],[MY] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX弯矩数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 10) = E_RcdSet.Fields("VX").Value
              Sheets("d_E").Cells(i + 2, 11) = E_RcdSet.Fields("MY").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取水平地震剪力Y、"弯矩X"
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
            StrSQL = "Select [Story],[VY],[MX] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " Max')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EY弯矩数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 14) = E_RcdSet.Fields("VY").Value
              Sheets("d_E").Cells(i + 2, 15) = E_RcdSet.Fields("MX").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取风荷载剪力X、弯矩Y，会出现负值，所以取了绝对值
          If OUTReader_Main.ComboBox_Wind_X.Text <> "" Then
            StrSQL = "Select [Story],[VX],[MY] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WX剪力数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 6) = Abs(E_RcdSet.Fields("VX").Value)
              Sheets("d_E").Cells(i + 2, 7) = Abs(E_RcdSet.Fields("MY").Value)
              E_RcdSet.MoveNext
            Next i
            End If
          End If
            
          '读取风荷载剪力Y、弯矩X
          If OUTReader_Main.ComboBox_Wind_Y.Text <> "" Then
            StrSQL = "Select [Story],[VY],[MX] From [Story Forces] Where [Location] IN ('Bottom') AND [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WY剪力数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 8) = Abs(E_RcdSet.Fields("VY").Value)
              Sheets("d_E").Cells(i + 2, 9) = Abs(E_RcdSet.Fields("MX").Value)
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          Call ETABS_DATA_CALC("Story Forces")
          
        '==================================================================================================================表 Story Shears '------------------------------------------for V9
        Case ("Story Shears")
          '读取水平地震剪力X、"弯矩Y"
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
            StrSQL = "Select [Story],[VX],[MY] From [Story Shears] Where [Loc] IN ('Bottom') AND [Load] IN  ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX弯矩数据不足！~"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 10) = E_RcdSet.Fields("VX").Value
              Sheets("d_E").Cells(i + 2, 11) = E_RcdSet.Fields("MY").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          '读取水平地震剪力Y、"弯矩X"
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
            StrSQL = "Select [Story],[VY],[MX] From [Story Shears] Where [Loc] IN ('Bottom') AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EY弯矩数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 14) = E_RcdSet.Fields("VY").Value
              Sheets("d_E").Cells(i + 2, 15) = E_RcdSet.Fields("MX").Value
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取风荷载剪力X、弯矩Y，会出现负值，所以取了绝对值
          If OUTReader_Main.ComboBox_Wind_X.Text <> "" Then
            StrSQL = "Select [Story],[VX],[MY] From [Story Shears] Where [Loc] IN ('Bottom') AND [Load] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WX剪力数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 6) = Abs(E_RcdSet.Fields("VX").Value)
              Sheets("d_E").Cells(i + 2, 7) = Abs(E_RcdSet.Fields("MY").Value)
              E_RcdSet.MoveNext
            Next i
            End If
          End If
            
          '读取风荷载剪力Y、弯矩X
          If OUTReader_Main.ComboBox_Wind_Y.Text <> "" Then
            StrSQL = "Select [Story],[VY],[MX] From [Story Shears] Where [Loc] IN ('Bottom') AND [Load] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "WY剪力数据不足！"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 8) = Abs(E_RcdSet.Fields("VY").Value)
              Sheets("d_E").Cells(i + 2, 9) = Abs(E_RcdSet.Fields("MX").Value)
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          Call ETABS_DATA_CALC("Story Forces")
          
        '==================================================================================================================表 Story Stiffness
       ' Case ("Story Stiffness")
'          '读取水平地震剪力X、刚度X、刚度比X
'          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
'            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
'                StrSQL = "Select [Story],[ShearX],[StiffX] From [Story Stiffness] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
'            ElseIf OUTReader_Main.Option_E_V9 Then
'                StrSQL = "Select [Story],[Shear-X],[Stiff-X] From [Story Stiffness] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
'            'Debug.Print StrSQL
'            Debug.Print "刚度比~"
'            End If
'
'                'StrSQL = "Select [Story],[ShearX],[StiffX] From [Story Stiffness] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
'            'Debug.Print StrSQL
'
'            Set E_RcdSet = New ADODB.Recordset
'            E_RcdSet.Open StrSQL, E_Connect, 3, 2
'
'            If E_RcdSet.RecordCount < Num_all Then
'              MsgBox "EX刚度数据不足！"
'            Else
'
'            For i = Num_all To 1 Step -1
'              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
'              'Sheets("d_E").Cells(i + 2, 10) = E_RcdSet.Fields("ShearX").Value
'
'                If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
'                    Sheets("d_E").Cells(i + 2, 4) = E_RcdSet.Fields("StiffX").Value
'                ElseIf OUTReader_Main.Option_E_V9 Then
'                    Sheets("d_E").Cells(i + 2, 4) = E_RcdSet.Fields("Stiff-X").Value
'                    'Debug.Print StrSQL
'                End If
'
'              'Sheets("d_E").Cells(i + 2, 4) = E_RcdSet.Fields("StiffX").Value
'              'Sheets("d_E").Cells(i + 2, 2) = E_RcdSet.Fields("Modifier").Value
'              E_RcdSet.MoveNext
'            Next i
'
'            End If
'          End If
'
'          '读取水平地震剪力Y、刚度Y、刚度比Y
'          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
'            If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
'                StrSQL = "Select [Story],[ShearY],[StiffY] From [Story Stiffness] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
'            ElseIf OUTReader_Main.Option_E_V9 Then
'                StrSQL = "Select [Story],[Shear-Y],[Stiff-Y] From [Story Stiffness] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
'            'Debug.Print StrSQL
'            End If
'
'                'StrSQL = "Select [Story],[ShearY],[StiffY] From [Story Stiffness] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
'            'Debug.Print StrSQL
'
'            Set E_RcdSet = New ADODB.Recordset
'            E_RcdSet.Open StrSQL, E_Connect, 3, 2
'
'            If E_RcdSet.RecordCount < Num_all Then
'              MsgBox "EY刚度数据不足！"
'            Else
'
'            For i = Num_all To 1 Step -1
'              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
'              'Sheets("d_E").Cells(i + 2, 14) = E_RcdSet.Fields("ShearY").Value
'
'                If OUTReader_Main.Option_E_V13 Then '-----------------------------添加判别
'                    Sheets("d_E").Cells(i + 2, 5) = E_RcdSet.Fields("StiffY").Value
'                ElseIf OUTReader_Main.Option_E_V9 Then
'                   Sheets("d_E").Cells(i + 2, 5) = E_RcdSet.Fields("Stiff-Y").Value
'                    'Debug.Print StrSQL
'                End If
'
'              'Sheets("d_E").Cells(i + 2, 5) = E_RcdSet.Fields("StiffY").Value
'              'Sheets("d_E").Cells(i + 2, 3) = E_RcdSet.Fields("Modifier").Value
'              E_RcdSet.MoveNext
'            Next i
'            End If
'          End If
'
'          '输出刚度比
'          For i = 1 To Num_all - 1
'          If Sheets("d_E").Cells(i + 3, 4) / Sheets("d_E").Cells(i + 3, 60) <> 0 And Sheets("d_E").Cells(i + 3, 5) / Sheets("d_E").Cells(i + 3, 60) Then
'             Sheets("d_E").Cells(i + 2, 2) = Sheets("d_E").Cells(i + 2, 4) * Sheets("d_E").Cells(i + 2, 60) / Sheets("d_E").Cells(i + 3, 4) / Sheets("d_E").Cells(i + 3, 60)
'             Sheets("d_E").Cells(i + 2, 3) = Sheets("d_E").Cells(i + 2, 5) * Sheets("d_E").Cells(i + 2, 60) / Sheets("d_E").Cells(i + 3, 5) / Sheets("d_E").Cells(i + 3, 60)
'          End If
'          Next
'          Sheets("d_E").Cells(Num_all + 2, 2) = 1
'          Sheets("d_E").Cells(Num_all + 2, 3) = 1
        Debug.Print "刚度比-------------"
          Call ETABS_DATA_CALC("Story Stiffness")
          
          
                    
        '==================================================================================================================表 Shear Gravity Ratios
        Case ("Shear Gravity Ratios")
          '读取剪重比X
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
            
                StrSQL = "Select [Story],[LambdaX],[LambdaMin] From [Shear Gravity Ratios] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX剪重比数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
                
              Sheets("d_E").Cells(i + 2, 12) = E_RcdSet.Fields("LambdaX").Value * 100
              Sheets("d_E").Cells(i + 2, 13) = E_RcdSet.Fields("LambdaMin").Value * 100
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          '读取剪重比Y
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
        
               StrSQL = "Select [Story],[LambdaY],[LambdaMin] From [Shear Gravity Ratios] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX剪重比数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
                
              Sheets("d_E").Cells(i + 2, 16) = E_RcdSet.Fields("LambdaY").Value * 100
              Sheets("d_E").Cells(i + 2, 17) = E_RcdSet.Fields("LambdaMin").Value * 100
              E_RcdSet.MoveNext
            Next i
            End If
          End If
          
          Call ETABS_DATA_CALC("Shear Gravity Ratios")
          
        '==================================================================================================================表 Shear/Gravity Ratios -----------------------------------------------------for V9
        Case ("Shear/Gravity Ratios")
          '读取剪重比X
          If OUTReader_Main.ComboBox_SPEC_X.Text <> "" Then
                StrSQL = "Select [Story],[Lambda-X],[LambdaMin] From [Shear/Gravity Ratios] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"

            
                'StrSQL = "Select [Story],[LambdaX],[LambdaMin] From [Shear Gravity Ratios] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX剪重比数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")

                    Sheets("d_E").Cells(i + 2, 12) = E_RcdSet.Fields("Lambda-X").Value * 100
                    Sheets("d_E").Cells(i + 2, 13) = E_RcdSet.Fields("LambdaMin").Value * 100

                
'              Sheets("d_E").Cells(i + 2, 12) = E_RcdSet.Fields("LambdaX").Value
'              Sheets("d_E").Cells(i + 2, 13) = E_RcdSet.Fields("LambdaMin").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          '读取剪重比Y
          If OUTReader_Main.ComboBox_SPEC_Y.Text <> "" Then
                StrSQL = "Select [Story],[Lambda-Y],[LambdaMin] From [Shear/Gravity Ratios] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"

            
               ' StrSQL = "Select [Story],[LambdaY],[LambdaMin] From [Shear Gravity Ratios] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
            'Debug.Print StrSQL
            
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "EX剪重比数据不足！"
              'Exit Sub'--------------------------------------------------------------------------------------------------------------------------------------------------此种直接退出sub太粗暴了……

            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")

                    Sheets("d_E").Cells(i + 2, 16) = E_RcdSet.Fields("Lambda-Y").Value * 100
                    Sheets("d_E").Cells(i + 2, 17) = E_RcdSet.Fields("LambdaMin").Value * 100
                
'              Sheets("d_E").Cells(i + 2, 16) = E_RcdSet.Fields("LambdaY").Value
'              Sheets("d_E").Cells(i + 2, 17) = E_RcdSet.Fields("LambdaMin").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          End If
          
          Call ETABS_DATA_CALC("Shear Gravity Ratios")
          
        '===============================================================================================================表 Mass Summary by Story
        Case ("Mass Summary by Story")
          Dim M_Base As String
          
          '读取楼层质量
          StrSQL = "Select [Story],[UX] From [Mass Summary by Story]" ' Order By [Mode] ASC"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "质量数据不足！"
            Else
            
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 54) = E_RcdSet.Fields("UX").Value
              E_RcdSet.MoveNext
            Next i
            End If
            
            '修正1层质量为1层+Base层
            M_Base = E_RcdSet.Fields("UX").Value
            'Debug.Print M_Base
            Sheets("d_E").Cells(3, 54) = Sheets("d_E").Cells(3, 54) + M_Base
          
          Call ETABS_DATA_CALC("Mass Summary by Story")
          

        '===============================================================================================================表Assembled Point Masses '-------------------for V9
        Case ("Assembled Point Masses")
          
          '读取楼层质量
          StrSQL = "Select [Story],[Point],[UX] From [Assembled Point Masses]"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount < Num_all Then
              MsgBox "质量数据不足！"
              'Debug.Print "楼层数" & E_RcdSet.RecordCount
            Else
                For i = E_RcdSet.RecordCount To 1 Step -1
                     If E_RcdSet.Fields("Point").Value = "All" Then
                        jj = extractNumberFromString2(E_RcdSet.Fields("Story").Value, 1)
                        If jj <> 0 Then
                            Sheets("d_E").Cells(jj + 2, 54) = E_RcdSet.Fields("UX").Value
                        ElseIf E_RcdSet.Fields("Story").Value = "BASE" Then
                            Sheets("d_E").Cells(Num_all + 3, 54) = E_RcdSet.Fields("UX").Value
                        End If

                     End If
                     E_RcdSet.MoveNext
                Next i
            
            End If
             '修正1层质量为1层 Base层
            a = Sheets("d_E").Cells(3, 54)
            Sheets("d_E").Cells(3, 54) = a + Sheets("d_E").Cells(Num_all + 3, 54)
            Sheets("d_E").Cells(Num_all + 3, 54) = ""

          
          Call ETABS_DATA_CALC("Mass Summary by Story")
          
          
        '===============================================================================================================表 Frame Shear Ratios In Dual Systems And Modifiers
        Case ("Frame Shear Ratios In Dual Systems And Modifiers")
          '读取X向数据
          StrSQL = "Select [Story],[Vo],[Vf],[Vf'],[Ratio] From [Frame Shear Ratios In Dual Systems And Modifiers] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "框架剪力数据不足！"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 48) = E_RcdSet.Fields("Vf").Value
              Sheets("d_E").Cells(i + 2, 49) = E_RcdSet.Fields("Vf").Value / E_RcdSet.Fields("Vo").Value
              Sheets("d_E").Cells(i + 2, 50) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
            
          '读取Y向数据
          StrSQL = "Select [Story],[Vo],[Vf],[Vf'],[Ratio] From [Frame Shear Ratios In Dual Systems And Modifiers] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "框架剪力数据不足！"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 51) = E_RcdSet.Fields("Vf").Value
              Sheets("d_E").Cells(i + 2, 52) = E_RcdSet.Fields("Vf").Value / E_RcdSet.Fields("Vo").Value
              Sheets("d_E").Cells(i + 2, 53) = E_RcdSet.Fields("Ratio").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
          
          
        '===============================================================================================================表 Frame Shear Ratios '-------------------for V9
        Case ("Frame Shear Ratios")
          
          '读取X向数据
          StrSQL = "Select [Story],[Vf],[Vf'],[Vf'/Vf] From [Frame Shear Ratios] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "框架剪力数据不足！"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 48) = E_RcdSet.Fields("Vf").Value
              'Sheets("d_E").Cells(i + 2, 49) = E_RcdSet.Fields("Vf").Value
              Sheets("d_E").Cells(i + 2, 50) = E_RcdSet.Fields("Vf'/Vf").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
            
          '读取Y向数据
          StrSQL = "Select [Story],[Vf],[Vf'],[Vf'/Vf] From [Frame Shear Ratios] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "框架剪力数据不足！"
            Else
            For i = Num_all To 1 Step -1
              'Debug.Print E_RcdSet.Fields("Story") & "," & E_RcdSet.Fields("Average")
              Sheets("d_E").Cells(i + 2, 51) = E_RcdSet.Fields("Vf").Value
              'Sheets("d_E").Cells(i + 2, 49) = E_RcdSet.Fields("Vf").Value
              Sheets("d_E").Cells(i + 2, 53) = E_RcdSet.Fields("Vf'/Vf").Value
              E_RcdSet.MoveNext
            Next i
            
            End If
            
         '===============================================================================================================表 Column Forces '-------------------for V9
        Case ("Column Forces")
          
          '读取VEX向数据
          For i = 1 To Num_all
            StrSQL = "Select [V3] From [Column Forces] Where [Loc] IN (0) AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & " ')"
            'Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox i & "层" & "缺少Column Forces数据！"
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 76) = WorksheetFunction.Sum(SS)
                Sheets("d_E").Cells(i + 2, 48).Formula = "=RC[28]"
                Sheets("d_E").Cells(i + 2, 49).Formula = "=RC[-1]/R3C80"
            End If
          Next i
         
          '读取VEY向数据
          For i = 1 To Num_all
            StrSQL1 = "Select [V2] From [Column Forces] Where [Loc] IN (0) AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " ')"
            Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL1, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               'MsgBox i & "层" & "缺少Column Forces(EY)数据！"
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 77) = WorksheetFunction.Sum(SS)
                Sheets("d_E").Cells(i + 2, 51).Formula = "=RC[26]"
                Sheets("d_E").Cells(i + 2, 52).Formula = "=RC[-1]/R3C81"
            End If
          Next i
            
           '读取MEX向数据
          For i = 1 To Num_all
            StrSQL = "Select [M2] From [Column Forces] Where [Loc] IN (0) AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XGD.Text & " ')"
            'Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               'MsgBox i & "层" & "缺少Column Forces(EX)数据！"
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 82) = WorksheetFunction.Sum(SS)
            End If
          Next i
         
          '读取MEY向数据
          For i = 1 To Num_all
            StrSQL1 = "Select [M3] From [Column Forces] Where [Loc] IN (0) AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YGD.Text & " ')"
            Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL1, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               'MsgBox i & "层" & "缺少Column Forces(EY)数据！"
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 83) = WorksheetFunction.Sum(SS)
            End If
          Next i

                     
            
         '===============================================================================================================表 Pier Forces '-------------------for V9
        Case ("Pier Forces")
          
          '读取VEX向数据
          For i = 1 To Num_all
            StrSQL = "Select [V2] From [Pier Forces] Where [Loc] IN ('Bottom') AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & " ')"
            'Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox i & "层" & "缺少Pier Forces数据！"
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 78) = WorksheetFunction.Sum(SS)
            End If
          Next i
         
          '读取VEY向数据
          For i = 1 To Num_all
            StrSQL1 = "Select [V3] From [Pier Forces] Where [Loc] IN ('Bottom') AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & " ')"
            Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL1, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 79) = WorksheetFunction.Sum(SS)
            End If
          Next i
            
           '读取MEX向数据
          For i = 1 To Num_all
            StrSQL = "Select [M3] From [Pier Forces] Where [Loc] IN ('Bottom') AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XGD.Text & " ')"
            'Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 84) = WorksheetFunction.Sum(SS)
            End If
          Next i
         
          '读取MEY向数据
          For i = 1 To Num_all
            StrSQL1 = "Select [M2] From [Pier Forces] Where [Loc] IN ('Bottom') AND [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YGD.Text & " ')"
            Debug.Print StrSQL
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL1, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 85) = WorksheetFunction.Sum(SS)
            End If
          Next i
          
          For i = 1 To Num_all
             Sheets("d_E").Cells(i + 2, 80).Formula = "=abs(RC[-4])+abs(RC[-2])"
             Sheets("d_E").Cells(i + 2, 81).Formula = "=abs(RC[-4])+abs(RC[-2])"
             Sheets("d_E").Cells(i + 2, 86).Formula = "=abs(RC[-4])+abs(RC[-2])"
             Sheets("d_E").Cells(i + 2, 87).Formula = "=abs(RC[-4])+abs(RC[-2])"
             Sheets("d_E").Cells(3, 88).Formula = "=ABS(RC[-6])/RC[-2]"
             Sheets("d_E").Cells(3, 89).Formula = "=ABS(RC[-6])/RC[-2]"
             Sheets("g_E").Cells(53, 5).Formula = "=d_E!R[-50]C[83]*100"
             Sheets("g_E").Cells(54, 5).Formula = "=d_E!R[-50]C[84]*100"
          Next i
        
            
        '===============================================================================================================表 Diaphragm CM Displacements '-------------------for V9
        Case ("Diaphragm CM Displacements")
          
          '读取EX向数据
          StrSQL = "Select [Story],[UX] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 18) = E_RcdSet.Fields("UX").Value
              E_RcdSet.MoveNext
            Next i
            End If
          '读取EX+向数据
          StrSQL = "Select [Story],[UX] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX+下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 19) = E_RcdSet.Fields("UX").Value
              E_RcdSet.MoveNext
            Next i
            End If
          '读取EX-向数据
          StrSQL = "Select [Story],[UX] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & "')"

          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2

            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EX-下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 20) = E_RcdSet.Fields("UX").Value
              E_RcdSet.MoveNext
            Next i
            End If
              '读取WX向数据
          StrSQL = "Select [Story],[UX] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "WX下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 21) = E_RcdSet.Fields("UX").Value
              E_RcdSet.MoveNext
            Next i
            End If
          '读取EY向数据
          StrSQL = "Select [Story],[UY] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 22) = E_RcdSet.Fields("UY").Value
              E_RcdSet.MoveNext
            Next i
            End If
          '读取EY+向数据
          StrSQL = "Select [Story],[UY] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY+下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 23) = E_RcdSet.Fields("UY").Value
              E_RcdSet.MoveNext
            Next i
            End If
          '读取EY-向数据
          StrSQL = "Select [Story],[UY] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "EY-下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 24) = E_RcdSet.Fields("UY").Value
              E_RcdSet.MoveNext
            Next i
            End If
              '读取WY向数据
          StrSQL = "Select [Story],[UY] From [Diaphragm CM Displacements] Where [Load] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "WY下位移数据不足！"
            Else
            For i = Num_all To 1 Step -1
              Sheets("d_E").Cells(i + 2, 25) = E_RcdSet.Fields("UY").Value
              E_RcdSet.MoveNext
            Next i
            End If
        Call ETABS_DATA_CALC("Diaphragm CM Displacements")
        
        
        

        '===============================================================================================================表 Frame Overturning Moments In Dual Systems
        Case ("Frame Overturning Moments In Dual Systems")
          '读取地震X向数据
          StrSQL = "Select [Story],[Total] From [Frame Overturning Moments In Dual Systems] Where [CaseCombo] IN ('~Static" & OUTReader_Main.ComboBox_SPEC_X.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "地震倾覆力矩数据不足！"
            Else
              
              E_RcdSet.MoveLast
              Sheets("g_E").Cells(50, 5) = E_RcdSet.Fields("Total").Value
            
            End If
            
          '读取地震Y向数据
          StrSQL = "Select [Story],[Total] From [Frame Overturning Moments In Dual Systems] Where [CaseCombo] IN ('~Static" & OUTReader_Main.ComboBox_SPEC_Y.Text & "')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "地震倾覆力矩数据不足！"
            Else

              E_RcdSet.MoveLast
              Sheets("g_E").Cells(51, 5) = E_RcdSet.Fields("Total").Value
              
            End If
          
          '读取风X向数据
          StrSQL = "Select [Story],[Total] From [Frame Overturning Moments In Dual Systems] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_X.Text & " 1')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "风倾覆力矩数据不足！"
            Else
              
              E_RcdSet.MoveLast
              Sheets("g_E").Cells(48, 5) = E_RcdSet.Fields("Total").Value
            
            End If
            
          '读取风Y向数据
          StrSQL = "Select [Story],[Total] From [Frame Overturning Moments In Dual Systems] Where [CaseCombo] IN ('" & OUTReader_Main.ComboBox_Wind_Y.Text & " 1')"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount <> Num_all Then
              MsgBox "风倾覆力矩数据不足！"
            Else

              E_RcdSet.MoveLast
              Sheets("g_E").Cells(49, 5) = E_RcdSet.Fields("Total").Value
              
            End If
        '===============================================================================================================表 Support Reactions
        Case ("Support Reactions")
          '读取恒载质量
          StrSQL = "Select [Story],[Load],[FZ] From [Support Reactions]"
          
          Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            
            If E_RcdSet.RecordCount = 0 Then
              MsgBox "缺少Support Reactions数据！"
            Else
            
                For i = E_RcdSet.RecordCount To 1 Step -1
                    If E_RcdSet.Fields("Story").Value = "Summation" And E_RcdSet.Fields("Load").Value = "DEAD" Then
                        Sheets("g_E").Cells(7, 5) = E_RcdSet.Fields("FZ").Value
                    ElseIf E_RcdSet.Fields("Story").Value = "Summation" And E_RcdSet.Fields("Load").Value = "LIVE" Then
                        Sheets("g_E").Cells(6, 5) = E_RcdSet.Fields("FZ").Value
                    End If
                  E_RcdSet.MoveNext
                Next i
            
            End If
          
        '===============================================================================================================表 Point Displacements
        Case ("Point Displacements")
          'Dim SS()
          
          '读取EX+
          For i = 1 To Num_all
            StrSQL = "Select [UX] From [Point Displacements] Where [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " ')"
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox "缺少Point Displacements(EX+)数据！"
               Exit Sub
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 70) = WorksheetFunction.Max(SS)
                Sheets("d_E").Cells(i + 2, 19) = WorksheetFunction.Average(SS)
            End If
        Next i

          
'          '读取EX+
'          For i = 1 To Num_all
'            StrSQL = "Select [UX] ,[STORY] From [Point Displacements] Where [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc.Text & " ')"
'            Set E_RcdSet = New ADODB.Recordset
'            E_RcdSet.Open StrSQL, E_Connect, 3, 2
'            Nu_NZ = E_RcdSet.RecordCount
'            If E_RcdSet.RecordCount = 0 Then
'               MsgBox "缺少Point Displacements(EX+)数据！"
'               Exit Sub
'            Else
'                ReDim SS(Nu_NZ - 1)
'                For j = 1 To E_RcdSet.RecordCount
'                    SS(j - 1) = E_RcdSet.Fields("UX").Value
'                    E_RcdSet.MoveNext
'                Next j
'                Sheets("d_E").Cells(i + 2, 70) = WorksheetFunction.Max(SS)
'            End If
'         Next i

          '读取EX-
          For i = 1 To Num_all
            StrSQL = "Select [UX] From [Point Displacements] Where [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_XEcc2.Text & " ')"
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox "缺少Point Displacements(EX-)数据！"
               Exit Sub
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 71) = WorksheetFunction.Max(SS)
                Sheets("d_E").Cells(i + 2, 20) = WorksheetFunction.Average(SS)
            End If
         Next i

          '读取EY+
          For i = 1 To Num_all
            StrSQL = "Select [UY]From [Point Displacements] Where [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc.Text & " ')"
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox "缺少Point Displacements(EY+)数据！"
               Exit Sub
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 72) = WorksheetFunction.Max(SS)
                Sheets("d_E").Cells(i + 2, 23) = WorksheetFunction.Average(SS)
            End If
         Next i
         
          '读取EY-
          For i = 1 To Num_all
            StrSQL = "Select [UY] From [Point Displacements] Where [STORY] IN ('" & "Story" & i & "')" & " AND [Load] IN ('" & OUTReader_Main.ComboBox_SPEC_YEcc2.Text & " ')"
            Set E_RcdSet = New ADODB.Recordset
            E_RcdSet.Open StrSQL, E_Connect, 3, 2
            Nu_NZ = E_RcdSet.RecordCount
            If E_RcdSet.RecordCount = 0 Then
               MsgBox "缺少Point Displacements(EY-)数据！"
               Exit Sub
            Else
                ReDim SS(Nu_NZ - 1)
                SS = E_RcdSet.GetRows
                Sheets("d_E").Cells(i + 2, 73) = WorksheetFunction.Max(SS)
                Sheets("d_E").Cells(i + 2, 24) = WorksheetFunction.Average(SS)
            End If
         Next i
         
         


            
            
            
            

          
          
        '==================================================================================================================
        Case Else
        
          
        End Select

  End If
  
  E_RxSchema.MoveNext
  
Loop

'关闭Access文件
E_RcdSet.Close
Set E_RcdSet = Nothing
E_RxSchema.Close
E_Connect.Close
Set E_Connect = Nothing




End Sub

