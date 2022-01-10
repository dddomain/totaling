Option Explicit

Public Sub makeSumFormula(execWs, formulas, ress, frstExecCell)

    Dim sumFormula As String
    Dim j As Long

    sumFormula = "=SUM(" & vbLf

    For j = 0 To ress
        sumFormula = sumFormula & formulas(j)
        If ress > j Then
            sumFormula = sumFormula & "," & vbLf
        End If
    Next j

    sumFormula = sumFormula & vbLf & ")"
    execWs.Range(frstExecCell) = sumFormula

End Sub
