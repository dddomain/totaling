Option Explicit

Function makeLinkColl(varsWs, ressWs, frstExecCell) As Collection

    Dim path As String: path = varsWs.Range("C2")
    Dim sht As String: sht = varsWs.Range("C4")
    Dim frstbk As String: frstbk = varsWs.Range("C5")
    Dim lstbk As String: lstbk = varsWs.Range("C6")
    Dim bk As String '= frstbk & resName & lstbk (関数の処理の中で代入)
    
    Dim ressTable As Variant: Set ressTable = ressWs.ListObjects(1)
    Dim linkColl As Collection: Set linkColl = New Collection
    
    Dim link As String
    Dim resName As String
    
    Dim i As Long
    'すべての回答のパスをCollectionに格納する
    For i = 1 To ressTable.ListColumns("Ress").DataBodyRange.Count
        resName = ressTable.ListColumns("Ress").DataBodyRange(i).Value
        bk = frstbk & resName & lstbk
        link = "'" & path & "[" & bk & "]" & sht & "'!" & frstExecCell
        linkColl.Add link
    Next i
    
    'return
    Set makeLinkColl = linkColl

End Function
