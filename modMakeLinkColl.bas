Option Explicit

'回答元テーブルの行数を確認

Function makeLinkColl(paramsTbl, ressTbl, frstExecCell) As Collection

    Dim path As String: path = paramsTbl.Range(2, 2)
    Dim sht As String: sht = paramsTbl.Range(4, 2)
    Dim frstbk As String: frstbk = paramsTbl.Range(5, 2)
    Dim lstbk As String: lstbk = paramsTbl.Range(6, 2)
    Dim bk As String '= frstbk & resName & lstbk (関数の処理の中で代入)
    
    Dim linkColl As Collection: Set linkColl = New Collection
    
    Dim link As String
    Dim resName As String
    
    Dim i As Long
    For i = 1 To ressTbl.ListColumns("Ress").DataBodyRange.Count
        resName = ressTbl.ListColumns("Ress").DataBodyRange(i).Value
        bk = frstbk & resName & lstbk
        link = "'" & path & "[" & bk & "]" & sht & "'!" & frstExecCell
        linkColl.Add link
    Next i
    
    'return
    Set makeLinkColl = linkColl

End Function
