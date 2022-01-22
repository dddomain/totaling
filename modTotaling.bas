Option Explicit

Sub Totaling()

    Dim execWs As Worksheet: Set execWs = ActiveSheet
    '各シートが存在するかをチェックしつつ代入する（=とりまとめシート上で実行しているかの確認）
    On Error GoTo ErrorHandler
    Dim varsWs As Worksheet: Set varsWs = Worksheets("変数（" & execWs.Name & "）")
    Dim ressWs As Worksheet: Set ressWs = Worksheets("回答元（" & execWs.Name & "）")
    Dim rngsWs As Worksheet: Set rngsWs = Worksheets("セル範囲（" & execWs.Name & "）")
    
    Dim frstExecCell As String: frstExecCell = varsWs.Range("C3")
    
    '和算式の再代入を行わない場合
    Dim rc As VbMsgBoxResult
    rc = MsgBox(frstExecCell & "に和算式を新規作成しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        GoTo NoMakeSumFormula
    End If
    
    Dim linkColl As Collection
    Set linkColl = makeLinkColl(varsWs, ressWs, frstExecCell)
    
    Dim sumformula As String
    sumformula = makeSumFormula(linkColl, frstExecCell)
    
    execWs.Range(frstExecCell) = sumformula

NoMakeSumFormula:

    rc = MsgBox(frstExecCell & "の式を全てのセルに代入しますか？", vbYesNo + vbQuestion)
    If rc = vbYes Then
        Call Spread(rngsWs, frstExecCell)
    End If
    
    '例外処理の前に脱出する
    Exit Sub
    
'例外処理
ErrorHandler:
        MsgBox Err.Description

End Sub
