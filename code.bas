Option Explicit

Sub totaling()

Dim ws As String
Dim path As String
Dim cellnum As String
Dim sht As String
Dim frstbk As String
Dim lstbk As String
Dim bk As String
Dim formula As String
Dim res_name As String
Dim formulas() As String

 
path = Worksheets("変数").Range("C2")
cellnum = Worksheets("変数").Range("C3")
sht = Worksheets("変数").Range("C4")
frstbk = Worksheets("変数").Range("C5")
lstbk = Worksheets("変数").Range("C6")
res_name = ""
bk = ""
formula = ""


'回答元テーブルの行数を確認

Dim i As Long
Dim lstrow As Long

lstrow = Worksheets("回答元").Cells(Rows.Count, 2).End(xlUp).Row


'すべての回答のパスを配列に格納する

ReDim formulas(lstrow - 2)

For i = 2 To lstrow
    res_name = Worksheets("回答元").Cells(i, 2).Value
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

    Range("A1").copy
    Range(totaling_cells).PasteSpecial xlPasteValues

End Sub

