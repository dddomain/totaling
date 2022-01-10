Option Explicit

Public Sub Spread(frstExecCell)

    'Range()の引数となるセル範囲をテーブルから取得してコレクションに格納する
    Dim rngWs As Worksheet: rngWs = Worksheets("セル範囲（" & ActiveSheet.Name & "）")
    Dim execRngTable As Variant: Set execRngTable = rngWs.ListObjects(1)
    Dim execRngColl As Variant: Set execRngColl = New Collection
    Dim gotRange As String
            
    Dim i As Long
    For i = 1 To execRngTable.ListColumns("Ranges").DataBodyRange.Count
        execRngColl.Add execRngTable.ListColumns("Ranges").DataBodyRange(i).Value
        gotRange = execRngColl(i)
        ActiveSheet.Range(gotRange).Select
        ActiveSheet.Range(frstExecCell).Copy ActiveSheet.Range(gotRange)
    Next i

End Sub
