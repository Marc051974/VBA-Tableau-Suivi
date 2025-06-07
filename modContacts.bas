Attribute VB_Name = "modContacts"
Sub MettreAJourListeContacts()
    Dim olApp As Object
    Dim olNS As Object
    Dim olFolder As Object
    Dim olContact As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    ' Initialiser Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If olApp Is Nothing Then
        MsgBox "Impossible de démarrer Outlook.", vbCritical
        Exit Sub
    End If

    ' Référencer le bon compte (contact@2mro.fr)
    Set olNS = olApp.GetNamespace("MAPI")
    On Error Resume Next
    Set olFolder = olNS.Folders("contact@2mro.fr").Folders("Contacts")
    On Error GoTo 0

    If olFolder Is Nothing Then
        MsgBox "Impossible d’accéder aux contacts de contact@2mro.fr", vbExclamation
        Exit Sub
    End If

    ' Supprimer ou créer la feuille ListeContacts
    Application.ScreenUpdating = False
    On Error Resume Next
    Worksheets("ListeContacts").Delete
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "ListeContacts"

    ' En-têtes
    ws.range("A1:D1").Value = Array("Société", "Prénom", "Téléphone", "Email")
    i = 2

    ' Parcourir les contacts
    For Each olContact In olFolder.Items
        If olContact.Class = 40 Then ' ContactItem uniquement
            ws.Cells(i, 1).Value = olContact.CompanyName
            ws.Cells(i, 2).Value = olContact.FirstName
            ws.Cells(i, 3).Value = olContact.MobileTelephoneNumber
            ws.Cells(i, 4).Value = olContact.Email1Address
            i = i + 1
        End If
    Next

    ws.Columns("A:D").AutoFit
    Application.ScreenUpdating = True

    MsgBox "Liste des contacts mise à jour.", vbInformation
End Sub

