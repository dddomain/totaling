Option Explicit

Sub Totaling()

    Dim execWs As Worksheet: Set execWs = ActiveSheet
    '各シートが存在するかをチェックしつつ代入する（=とりまとめシート上で実行しているかの確認）
    On Error GoTo ErrorHandler
    Dim varsWs As Worksheet: Set varsWs = Worksheets("変数（" & execWs.Name & "）")

    Dim paramsTbl As Variant: Set paramsTbl = varsWs.ListObjects("params")
    Dim execRngsTbl As Variant: Set execRngsTbl = varsWs.ListObjects("execRngs")
    Dim ressTbl As Variant: Set ressTbl = varsWs.ListObjects("ress")
    
    Dim frstExecCell As String: frstExecCell = paramsTbl.Range(3, 2)
    
    '和算式の再代入を行わない場合
    Dim rc As VbMsgBoxResult
    rc = MsgBox(frstExecCell & "に和算式を新規作成しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        GoTo NoMakeSumFormula
    End If
    
    Dim linkColl As Collection
    Set linkColl = makeLinkColl(paramsTbl, ressTbl, frstExecCell)
    
    Dim sumformula As String
    sumformula = makeSumFormula(linkColl, frstExecCell)
    
    execWs.Range(frstExecCell) = sumformula

NoMakeSumFormula:

    rc = MsgBox(frstExecCell & "の式を全てのセルに代入しますか？", vbYesNo + vbQuestion)
    If rc = vbYes Then
        Call Spread(execRngsTbl, frstExecCell)
    End If
    
    '例外処理の前に脱出する
    Exit Sub
    
'例外処理
ErrorHandler:
        MsgBox Err.Description

End Sub

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

Function makeSumFormula(linkColl, frstExecCell) As String

    Dim sumformula As String
    sumformula = "=SUM(" & vbLf

    Dim j As Long
    For j = 1 To linkColl.Count
        sumformula = sumformula & linkColl(j)
        If linkColl.Count > j Then
            sumformula = sumformula & "," & vbLf
        End If
    Next j
    sumformula = sumformula & vbLf & ")"

    'return
    makeSumFormula = sumformula

End Function

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
