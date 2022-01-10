Option Explicit

Function makeSumFormula(linkColl, frstExecCell) As String

    Dim sumformula As String
    Dim j As Long
    
    sumformula = "=SUM(" & vbLf
    
    For j = 1 To linkColl.Count
        sumformula = sumformula & linkColl(j)
        If linkColl.Count > j Then
            sumformula = sumformula & "," & vbLf
        End If
    Next j
    
    sumformula = sumformula & vbLf & ")"
    makeSumFormula = sumformula

End Function
