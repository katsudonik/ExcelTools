Attribute VB_Name = "getRange"
'    NOTE: Usage:
'    endRow = getRange("endRow")
'    endColumn = getRange("endColumn")
Function getRange() As Collection
    Dim ExistsRange As New Collection
    Dim oRange   As Range
    Dim endRow As Integer
    Dim endColumn As Integer

    Set oRange = ActiveSheet.UsedRange
    ExistsRange.Add oRange.Row, "firstRow"
    endRow = oRange.Row + oRange.Rows.Count - 1
    ExistsRange.Add endRow, "endRow"
    ExistsRange.Add oRange.Column, "firstColumn"
    endColumn = oRange.Column + oRange.Columns.Count - 1
    ExistsRange.Add endColumn, "endColumn"
    
    Set getRange = ExistsRange
End Function


