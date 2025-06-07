Attribute VB_Name = "Module2"
Sub GenererListeRelance()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Suivi") '
    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    
    Dim i As Long
    Dim listeRelance As String
    listeRelance = "Clients à relancer :" & vbCrLf
    Dim compteRelances As Integer
    compteRelances = 0
    
    For i = 2 To derniereLigne
        If IsDate(ws.Cells(i, 2).Value) Then
         If ws.Cells(i, 2).Value = DateSerial(2024, 3, 25) And ws.Cells(i, 5).Value = "En attente" Then

                listeRelance = listeRelance & ws.Cells(i, 3).Value & vbCrLf
                compteRelances = compteRelances + 1
            End If
        End If
    Next i
    
    If compteRelances = 0 Then
        listeRelance = "Aucun client à relancer pour la période sélectionnée."
    End If
    
    MsgBox listeRelance
End Sub
