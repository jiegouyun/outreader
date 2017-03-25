Attribute VB_Name = "Etabs_Load_Cases"

'******************************************************************************
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'****                        ETABS荷载工况选择                             ****
'****                                                                      ****
'****                                                                      ****
'****                                                                      ****
'******************************************************************************
'******************************************************************************

'////////////////////////////////////////////////////////////////////////////

'更新时间:2014/3/4/12：24
'更新内容:
'1.ETABS荷载工况选择，针对从MDB文件中读取


Public Sub Etabs_Load_Case(MDB_Path As String)

'===================================================================================================读取荷载工况

'定义变量
Dim E_Connect As New ADODB.Connection
Dim E_RcdSet As New ADODB.Recordset
Dim StrSQL As String
Dim Static_Cases(), Spec_Cases(), Hist_Cases(), GDSpec_Cases() As String
Dim N_Static, N_Spec, N_Hist As Long
Dim GD As Integer '添加个判断是否自行添加规定水平力选项

GD = 0

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


If OUTReader_Main.Option_E_V13 Then

    '读取静力工况，包括风荷载
    StrSQL = "Select Distinct [Name] From [Load Cases - Summary] Where [Type] Like '%Static'AND Not([Name] Like '~%')"
    
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    N_Static = E_RcdSet.RecordCount
    ReDim Static_Cases(N_Static - 1)
    For i = 0 To E_RcdSet.RecordCount - 1
      Static_Cases(i) = E_RcdSet.Fields("Name").Value
      'Debug.Print Static_Cases(i)
      E_RcdSet.MoveNext
    Next i
    
    '读取反应谱工况
    StrSQL = "Select Distinct [Name] From [Load Cases - Summary] Where [Type] Like '%Spectrum'"
    
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    N_Spec = E_RcdSet.RecordCount
    If N_Spec = 0 Then
        MsgBox "缺少反应谱工况！"
    Else
        ReDim Spec_Cases(N_Static - 1)
        For i = 0 To E_RcdSet.RecordCount - 1
          Spec_Cases(i) = E_RcdSet.Fields("Name").Value
          'Debug.Print Spec_Cases(i)
          E_RcdSet.MoveNext
        Next i
    End If
    
'    '读取时程工况
'    StrSQL = "Select Distinct [Name] From [Load Cases - Summary] Where [Type] Like '%History'"
'
'    Set E_RcdSet = New ADODB.Recordset
'    E_RcdSet.Open StrSQL, E_Connect, 3, 2
'
'    'N_Hist = E_RcdSet.RecordCount
'    If N_Hist = 0 Then '-------------------------------------------------------------------------------------------------------------------添加判别
'        MsgBox "缺少时程工况"
'    Else
'        ReDim Hist_Cases(E_RcdSet.RecordCount - 1)
'        For i = 0 To E_RcdSet.RecordCount - 1
'          Hist_Cases(i) = E_RcdSet.Fields("Name").Value
'          'Debug.Print Hist_Cases(i)
'          E_RcdSet.MoveNext
'        Next i
'    End If
    
ElseIf OUTReader_Main.Option_E_V9 Then

   '读取静力工况，包括风荷载
    StrSQL = "Select Distinct [Case] From [Static Load Cases] Where [Type] Like '%WIND'"
    
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    N_Static = E_RcdSet.RecordCount
    
    If N_Static = 0 Then
        MsgBox "荷载工况没有选择!"
        Exit Sub
    Else
        ReDim Static_Cases(N_Static - 1)
    End If
    
    For i = 0 To E_RcdSet.RecordCount - 1
      Static_Cases(i) = E_RcdSet.Fields("Case").Value
      'Debug.Print Static_Cases(i)
      E_RcdSet.MoveNext
    Next i
    
   '读取静力工况，规定水平力
    StrSQL = "Select Distinct [Case] From [Static Load Cases] Where [Type] Like '%QUAKE'"
    
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    N_Static = E_RcdSet.RecordCount
    
    If N_Static = 0 Then
        MsgBox "规定水平力工况没有选择!"
        'Exit Sub
    Else
        GD = 1
        ReDim GDSpec_Cases(N_Static - 1)
    End If
    
    For i = 0 To E_RcdSet.RecordCount - 1
      GDSpec_Cases(i) = E_RcdSet.Fields("Case").Value
      'Debug.Print Static_Cases(i)
      E_RcdSet.MoveNext
    Next i
    
    '读取反应谱工况
    StrSQL = "Select Distinct [Case] From [Response Spectrum Cases]"
    
    Set E_RcdSet = New ADODB.Recordset
    E_RcdSet.Open StrSQL, E_Connect, 3, 2
    
    N_Spec = E_RcdSet.RecordCount
    If N_Spec = 0 Then
        MsgBox "缺少反应谱工况！"
    Else
        ReDim Spec_Cases(N_Spec - 1)
        For i = 0 To E_RcdSet.RecordCount - 1
          Spec_Cases(i) = E_RcdSet.Fields("Case").Value
          'Debug.Print Spec_Cases(i)
          E_RcdSet.MoveNext
        Next i
    End If
    
'    '读取时程工况
'    StrSQL = "Select Distinct [Name] From [Load Cases - Summary] Where [Type] Like '%History'"
'
'    Set E_RcdSet = New ADODB.Recordset
'    E_RcdSet.Open StrSQL, E_Connect, 3, 2
'
'    N_Hist = E_RcdSet.RecordCount
'    If N_Hist = 0 Then '-------------------------------------------------------------------------------------------------------------------添加判别
'        MsgBox "缺少时程工况"
'    Else
'        ReDim Hist_Cases(E_RcdSet.RecordCount - 1)
'        For i = 0 To E_RcdSet.RecordCount - 1
'          Hist_Cases(i) = E_RcdSet.Fields("Name").Value
'          'Debug.Print Hist_Cases(i)
'          E_RcdSet.MoveNext
'        Next i
'    End If
    
End If

'关闭Access文件
E_RcdSet.Close
Set E_RcdSet = Nothing
E_Connect.Close
Set E_Connect = Nothing

'===================================================================================================定义复合框列表

If GD = 0 Then
'反应谱工况
OUTReader_Main.ComboBox_SPEC_X.List = Spec_Cases
OUTReader_Main.ComboBox_SPEC_XEcc.List = Spec_Cases
OUTReader_Main.ComboBox_SPEC_Y.List = Spec_Cases
OUTReader_Main.ComboBox_SPEC_YEcc.List = Spec_Cases

ElseIf GD = 1 Then

'反应谱工况
OUTReader_Main.ComboBox_SPEC_X.List = Spec_Cases
OUTReader_Main.ComboBox_SPEC_XEcc.List = GDSpec_Cases
OUTReader_Main.ComboBox_SPEC_XEcc2.List = GDSpec_Cases
OUTReader_Main.ComboBox_SPEC_XGD.List = GDSpec_Cases
OUTReader_Main.ComboBox_SPEC_Y.List = Spec_Cases
OUTReader_Main.ComboBox_SPEC_YEcc.List = GDSpec_Cases
OUTReader_Main.ComboBox_SPEC_YEcc2.List = GDSpec_Cases
OUTReader_Main.ComboBox_SPEC_YGD.List = GDSpec_Cases

End If

'风荷载工况
OUTReader_Main.ComboBox_Wind_X.List = Static_Cases
OUTReader_Main.ComboBox_Wind_Y.List = Static_Cases

'时程工况
If N_Hist > 1 Then '-------------------------------------------------------------------------------------------------------------------添加判别
OUTReader_Main.ListBox_TH_X.List = Hist_Cases
OUTReader_Main.ListBox_TH_Y.List = Hist_Cases
End If

End Sub
