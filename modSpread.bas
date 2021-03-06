Option Explicit

Function Spread(execRngsTbl, frstExecCell)
    
    'Range()の引数となるセル範囲をテーブルから取得してコレクションに格納する
    Dim execRngColl As Variant: Set execRngColl = New Collection
    Dim gotRange As String
    Dim i As Long
    For i = 1 To execRngsTbl.ListColumns("Ranges").DataBodyRange.Count
        execRngColl.Add execRngsTbl.ListColumns("Ranges").DataBodyRange(i).Value
        gotRange = execRngColl(i)
        ActiveSheet.Range(gotRange).Select
        ActiveSheet.Range(frstExecCell).Copy ActiveSheet.Range(gotRange)
    Next i

End Function
