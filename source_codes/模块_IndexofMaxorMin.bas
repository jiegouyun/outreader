Attribute VB_Name = "Ä£¿é_IndexofMaxorMin"
Option Explicit
Function IndexMinofRangeM(index_Range As range)
Dim R, C As Integer
Dim Min As Variant
Min = WorksheetFunction.Min(index_Range)
R = index_Range.Find(Min, LookIn:=xlValues).Row
C = index_Range.Find(Min, LookIn:=xlValues).Columns
IndexMinofRangeM = Array(Min, R, C)
End Function


Function IndexMaxofRangeM(index_Range As range)
Dim R, C As Integer
Dim Max As Variant
Max = WorksheetFunction.Max(index_Range)
R = index_Range.Find(Max, LookIn:=xlValues).Row
C = index_Range.Find(Max, LookIn:=xlValues).Columns
IndexMaxofRangeM = Array(Max, R, C)
End Function


