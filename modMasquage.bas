Attribute VB_Name = "modMasquage"
Sub MasquerLignesRefusees()
    Dim ws As Worksheet
    Dim range As range
    Dim row As range
    Dim cell As range
    
    Set ws = ThisWorkbook.Sheets("Suivi")
    Set range = ws.range("L2:L1000")
    
    For Each row In range.Rows
        Set cell = row.Cells(1, 1)
        If cell.Value = "Refusé" Or cell.Value = "Accepté" Then
            row.EntireRow.Hidden = True
        Else
            row.EntireRow.Hidden = False
        End If
    Next row
End Sub
