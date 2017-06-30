Attribute VB_Name = "getRange"
'    Set ExistsRange = UsedRangeSample
'    endRow = ExistsRange("endRow")
'
'    Set ExistsRange = UsedRangeSample
'    endColumn = ExistsRange("endColumn")

'Excel最終行の取得
Function getRange() As Collection
    Dim ExistsRange As New Collection
    Dim oRange   As Range
    Dim endRow As Integer
    Dim endColumn As Integer

    'UsedRangeでデータの範囲を自動的に求めます
    Set oRange = ActiveSheet.UsedRange

    '範囲から、上下の行番号と左右の列番号を求めます
    '開始行
    ExistsRange.Add oRange.Row, "firstRow"
    '最終行
    endRow = oRange.Row + oRange.Rows.Count - 1
    ExistsRange.Add endRow, "endRow"
    
    '範囲から、左右の列番号を求めます
    '開始列
    ExistsRange.Add oRange.Column, "firstColumn"
    '最終列
    endColumn = oRange.Column + oRange.Columns.Count - 1
    ExistsRange.Add endColumn, "endColumn"
    
    Set UsedRangeSample = ExistsRange
End Function


