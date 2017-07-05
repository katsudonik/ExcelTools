Attribute VB_Name = "Timer"
'開始時間の保持用変数
Private sttime As Long

'タイマー開始
Private Sub ExTimerStart()
    Dim ln As Long
    Dim lm As Long
    
    Do
        '経過時間を算出
        ln = Timer - sttime
        lm = Int(ln / 60)
        '○分○○秒で表示
        Range("D5") = lm & "分" & Format(ln - lm * 60, "00") & "秒"
        DoEvents
        If ln >= 180 Then
            Beep
            MsgBox "３分経過しました。さあ、食べてください。"
            Exit Do
        End If
    Loop
    
End Sub

'スタートボタン
Private Sub CommandButton1_Click()
    Range("D5") = ""
    '開始時間
    sttime = Timer
    'タイマー開始
    ExTimerStart
End Sub

