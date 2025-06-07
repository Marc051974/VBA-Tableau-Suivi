VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Call Relancer_Offres_Prix

End Sub
Public Sub TraiterEntreprise(ByVal Target As range)
    Dim wsContacts As Worksheet
    Dim entreprise As String
    Dim prenoms As Collection
    Dim cell As range
    Dim i As Long

    If Target.Column = 1 And Target.row > 1 Then
        entreprise = LCase(Trim(Target.Value))
        If entreprise = "" Then Exit Sub

        Set wsContacts = ThisWorkbook.Sheets("ListeContacts")
        Set prenoms = New Collection

        ' Recherche des prénoms associés
        For Each cell In wsContacts.range("A2:A" & wsContacts.Cells(wsContacts.Rows.Count, 1).End(xlUp).row)
            If LCase(Trim(cell.Value)) = entreprise Then
                On Error Resume Next
                prenoms.Add cell.Offset(0, 1).Value, CStr(cell.Offset(0, 1).Value)
                On Error GoTo 0
            End If
        Next cell

        If prenoms.Count = 0 Then
            MsgBox "Aucun contact trouvé pour """ & Target.Value & """", vbExclamation
            Exit Sub
        End If

        ' Créer la liste déroulante en B
        Dim liste As String
        For i = 1 To prenoms.Count
            liste = liste & prenoms(i) & ","
        Next i
        liste = Left(liste, Len(liste) - 1)

        With Cells(Target.row, 2).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:=liste
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As range)
    If Target.Column = 2 And Target.row > 1 Then
        Dim wsContacts As Worksheet
        Set wsContacts = ThisWorkbook.Sheets("ListeContacts")

        Dim nomEntreprise As String
        nomEntreprise = Cells(Target.row, 1).Value
        Dim prenom As String
        prenom = Target.Value

        Dim cell As range
        For Each cell In wsContacts.range("A2:A" & wsContacts.Cells(wsContacts.Rows.Count, 1).End(xlUp).row)
            If cell.Value = nomEntreprise And cell.Offset(0, 1).Value = prenom Then
                Cells(Target.row, 4).Value = cell.Offset(0, 2).Value ' Portable
                Cells(Target.row, 5).Value = cell.Offset(0, 3).Value ' Email
                Exit For
            End If
        Next cell
    End If
End Sub


