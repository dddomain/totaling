Option Explicit

Sub totaling()
Dim varWs As Worksheet: Set varWs = Worksheets("変数（とりまとめ）")
Dim path As String: path = varWs.Range("C2")
Dim cellnum As String: cellnum = varWs.Range("C3")
Dim sht As String: sht = varWs.Range("C4")
Dim frstbk As String: frstbk = varWs.Range("C5")
Dim lstbk As String: lstbk = varWs.Range("C6")
Dim bk As String
Dim formula As String
Dim res_name As String
Dim formulas() As String

'回答元テーブルの行数を確認

Dim i As Long
Dim lstrow As Long

lstrow = Worksheets("回答元（とりまとめ）").Cells(Rows.Count, 2).End(xlUp).Row

'すべての回答のパスを配列に格納する
ReDim formulas(lstrow - 2)
For i = 2 To lstrow
    res_name = Worksheets("回答元（とりまとめ）").Cells(i, 2).Value
    bk = frstbk & res_name & lstbk
    formula = "'" & path & "[" & bk & "]" & sht & "'!" & cellnum
    formulas(i - 2) = formula
Next i

Dim ress As Long
ress = i - 3

Call make_sum_formula(formulas, ress, cellnum)
Call copy_formula

End Sub

Private Sub make_sum_formula(formulas, ress, cellnum)

Dim sum_formula As String
Dim j As Long

sum_formula = "=SUM(" & vbLf

For j = 0 To ress
    sum_formula = sum_formula & formulas(j)
    If ress > j Then
        sum_formula = sum_formula & "," & vbLf
    End If
Next j

sum_formula = sum_formula & vbLf & ")"
MsgBox sum_formula
Worksheets(1).Range(cellnum) = sum_formula

End Sub

Private Sub copy_formula()

    Dim totaling_cells As String
    totaling_cells = "A2, B1:B2"

    Range("A1").Copy
    Range(totaling_cells).PasteSpecial xlPasteValues

End Sub


