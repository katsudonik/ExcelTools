Attribute VB_Name = "getRange"
'    Set ExistsRange = UsedRangeSample
'    endRow = ExistsRange("endRow")
'
'    Set ExistsRange = UsedRangeSample
'    endColumn = ExistsRange("endColumn")

'Excel�ŏI�s�̎擾
Function getRange() As Collection
    Dim ExistsRange As New Collection
    Dim oRange   As Range
    Dim endRow As Integer
    Dim endColumn As Integer

    'UsedRange�Ńf�[�^�͈̔͂������I�ɋ��߂܂�
    Set oRange = ActiveSheet.UsedRange

    '�͈͂���A�㉺�̍s�ԍ��ƍ��E�̗�ԍ������߂܂�
    '�J�n�s
    ExistsRange.Add oRange.Row, "firstRow"
    '�ŏI�s
    endRow = oRange.Row + oRange.Rows.Count - 1
    ExistsRange.Add endRow, "endRow"
    
    '�͈͂���A���E�̗�ԍ������߂܂�
    '�J�n��
    ExistsRange.Add oRange.Column, "firstColumn"
    '�ŏI��
    endColumn = oRange.Column + oRange.Columns.Count - 1
    ExistsRange.Add endColumn, "endColumn"
    
    Set UsedRangeSample = ExistsRange
End Function


