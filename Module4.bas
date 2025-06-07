Attribute VB_Name = "Module4"
Sub RechercherEntreprise()
    Dim texteRecherche As String

    ' Forcer l’activation de la feuille Suivi si besoin
    If ActiveSheet.Name <> "Suivi" Then
        Sheets("Suivi").Activate
    End If

    ' Si la cellule active est vide ou n’est pas en colonne A, demander à l’utilisateur
    If ActiveCell.Column <> 1 Or Trim(ActiveCell.Value) = "" Then
        texteRecherche = InputBox("Tapez quelques lettres du nom de l’entreprise à rechercher :", "Recherche entreprise")
        If texteRecherche = "" Then Exit Sub ' Annulation ou champ vide
        Sheets("Suivi").range("A1").Value = texteRecherche
        Sheets("Suivi").range("A1").Select ' Active une cellule temporaire avec le texte
    End If

    ' Lancer le formulaire
    frmRechercheEntreprise.Show
End Sub

