Option Explicit

Sub test()

Dim pass As String
pass = "D:\15　消防庁・内閣府からの照会\Ｒ３年度\消防庁より\211224〆防災拠点となる公共施設等の耐震化推進状況調査等について\03庁内各課から回収\"

Dim bookname As String
Dim sheetname As String

bookname = "07+（今治市）【様式1／様式2／様式3-1／様式3-2／様式3-3】公共施設等耐震化（都道府県／市町村)+.xlsx"
sheetname = "様式２（都道府県）"

'Worksheets(2).Range("E9").Value = "=Sum(1 + 2)"

MsgBox "='" & pass & "[" & bookname & "]!" & sheetname & "!'Range('E9')"

'Worksheets(2).Range("E9").Value = "='" & pass & "[" & bookname & "]" & sheetname & "'!Range(E9)"
Worksheets(2).Range("E9").Value = "='" & pass & "[" & bookname & "]" & sheetname & "'!E9"

End Sub
