Attribute VB_Name = "Timer"
'�J�n���Ԃ̕ێ��p�ϐ�
Private sttime As Long

'�^�C�}�[�J�n
Private Sub ExTimerStart()
    Dim ln As Long
    Dim lm As Long
    
    Do
        '�o�ߎ��Ԃ��Z�o
        ln = Timer - sttime
        lm = Int(ln / 60)
        '���������b�ŕ\��
        Range("D5") = lm & "��" & Format(ln - lm * 60, "00") & "�b"
        DoEvents
        If ln >= 180 Then
            Beep
            MsgBox "�R���o�߂��܂����B�����A�H�ׂĂ��������B"
            Exit Do
        End If
    Loop
    
End Sub

'�X�^�[�g�{�^��
Private Sub CommandButton1_Click()
    Range("D5") = ""
    '�J�n����
    sttime = Timer
    '�^�C�}�[�J�n
    ExTimerStart
End Sub

