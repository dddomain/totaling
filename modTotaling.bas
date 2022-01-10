Option Explicit

Sub Totaling()

    Dim execWs As Worksheet: Set execWs = ActiveSheet
    '各シートが存在するかをチェックしつつ代入する（=とりまとめシート上で実行しているかの確認）
    On Error GoTo ErrorHandler
    Dim varsWs As Worksheet: Set varsWs = Worksheets("変数（" & execWs.Name & "）")
    Dim ressWs As Worksheet: Set ressWs = Worksheets("回答元（" & execWs.Name & "）")

    Dim path As String: path = varsWs.Range("C2")
    Dim frstExecCell As String: frstExecCell = varsWs.Range("C3")
    Dim sht As String: sht = varsWs.Range("C4")
    Dim frstbk As String: frstbk = varsWs.Range("C5")
    Dim lstbk As String: lstbk = varsWs.Range("C6")
    Dim bk As String '関数の処理の中で代入

    Dim formula As String
    Dim resName As String '関数の処理の中で代入
    Dim formulas() As String

    '回答元テーブルの行数を確認

    Dim i As Long
    Dim lstrow As Long

    lstrow = ressWs.Cells(Rows.Count, 2).End(xlUp).Row

    'すべての回答のパスを配列に格納する
    ReDim formulas(lstrow - 2)
    For i = 2 To lstrow
        resName = ressWs.Cells(i, 2).Value
        bk = frstbk & resName & lstbk
        formula = "'" & path & "[" & bk & "]" & sht & "'!" & frstExecCell
        formulas(i - 2) = formula
    Next i

    Dim ress As Long
    ress = i - 3

    Call makeSumFormula(execWs, formulas, ress, frstExecCell)


    Dim rc As VbMsgBoxResult
    rc = MsgBox(frstExecCell & "で作成した式を全てのセルに代入しますか？", vbYesNo + vbQuestion)
    If rc = vbYes Then
        Call Spread(frstExecCell)
    End If

    Exit Sub

    '例外処理
    ErrorHandler:
        MsgBox "とりまとめを行うシートで実行してください。"

End Sub

