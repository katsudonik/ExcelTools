Attribute VB_Name = "progressBar"
Sub progressBarSample()
    Dim i As Long
    endCnt = 500
    For i = 1 To endCnt
            Call progressBar(i, endCnt)
    Next i
    Application.StatusBar = False
End Sub

Function progressBar(Cnt, endCnt)
    Application.StatusBar = "èàóùíÜ..." & String(Int(Cnt / endCnt * 100), "Å°") & String(Int(100 - Cnt / endCnt * 100), "Å†") & Int(Cnt / endCnt * 100) & "%"
    If Cnt = endCnt Then
        Application.StatusBar = False
    End If
End Function

