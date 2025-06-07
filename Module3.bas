Attribute VB_Name = "Module3"
Option Explicit

Public Sub Relancer_Offres_Prix()
    Dim ws            As Worksheet
    Dim lastRow       As Long, i As Long, cnt As Long
    Dim vEnvoi        As Variant, vDelai As Variant, vRel As Variant
    Dim dEnvoi        As Date, dDelai As Date, dRel As Date
    Dim pronom        As String, client As String, email As String
    Dim usine         As String, refChantier As String, statut As String
    Dim pieceJointe   As String, corpsMail As String
    Dim OutlookApp    As Object, OutlookMail As Object

    Set ws = ThisWorkbook.Sheets("Suivi")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    cnt = 0

    ' Initialiser Outlook
    On Error Resume Next
    Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        MsgBox "Erreur : Outlook non disponible.", vbCritical
        Exit Sub
    End If

    For i = 2 To lastRow
        ' — Lecture brute —
        pronom = LCase$(Trim(ws.Cells(i, 3).Value))         ' Col C
        client = Trim(ws.Cells(i, 2).Value)                 ' Col B
        email = Trim(ws.Cells(i, 5).Value)                  ' Col E
        usine = Trim(ws.Cells(i, 6).Value)                  ' Col F
        refChantier = Trim(ws.Cells(i, 7).Value)            ' Col G
        statut = LCase$(Trim(ws.Cells(i, 12).Value))        ' Col L
        pieceJointe = Trim(ws.Cells(i, 14).Value)           ' Col N

        ' — Variants avant conversion —
        vEnvoi = ws.Cells(i, 8).Value    ' Col H
        vDelai = ws.Cells(i, 10).Value   ' Col J
        vRel = ws.Cells(i, 11).Value     ' Col K

        ' — Conversion datée —
        If IsDate(vEnvoi) Then
            dEnvoi = CDate(vEnvoi)
        Else
            GoTo SkipRow
        End If
        If IsDate(vDelai) Then
            dDelai = CDate(vDelai)
        Else
            dDelai = 0
        End If
        If IsDate(vRel) Then
            dRel = CDate(vRel)
        Else
            dRel = 0
        End If

        ' — 1) Délai souhaité (si renseigné) —
        If dDelai <> 0 And dDelai > Date Then GoTo SkipRow

        ' — 2) Date d’envoi = 60 jours —
        If DateDiff("d", dEnvoi, Date) < 60 Then GoTo SkipRow

        ' — 3) Dernière relance vide OU = 60 jours —
        If dRel <> 0 Then
            If DateDiff("d", dRel, Date) < 60 Then GoTo SkipRow
        End If

        ' — 4) Statut = “en attente” —
        If statut <> "en attente" Then GoTo SkipRow

        ' === on passe tous les filtres ===
        cnt = cnt + 1

        ' — Construction du corps —
        corpsMail = "<p>Bonjour " & client & ",</p>"
        If pronom = "tu" Then
            corpsMail = corpsMail & _
                "<p>Je reviens vers toi concernant notre offre <b>" & usine & _
                "</b> envoyée le <b>" & Format(dEnvoi, "dd/mm/yyyy") & _
                "</b> pour le dossier <b>" & refChantier & "</b>.</p>" & _
                "<p>Peux-tu m'indiquer l'état d'avancement de ton projet ?</p>" & _
                "<div>Merci pour ton retour.</div>"
        Else
            corpsMail = corpsMail & _
                "<p>Je reviens vers vous concernant notre offre <b>" & usine & _
                "</b> envoyée le <b>" & Format(dEnvoi, "dd/mm/yyyy") & _
                "</b> pour le dossier <b>" & refChantier & "</b>.</p>" & _
                "<p>Pouvez-vous m'indiquer l'état d'avancement de votre projet ?</p>" & _
                "<div>Merci pour votre retour.</div>"
        End If
        corpsMail = corpsMail & _
            "<ul><li><b>Projet validé</b></li>" & _
                   "<li><b>Date de relance souhaitée</b></li>" & _
                   "<li><b>Offre non retenue</b></li></ul>"

        ' — Création et ouverture du mail —
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = email
            .Subject = "Suivi de votre offre - " & refChantier
            .Display                                   ' Ouvre le mail avec signature
            Application.Wait Now + TimeValue("00:00:02")
            .HTMLBody = corpsMail & .HTMLBody          ' Insère le corps au-dessus
            
            ' — Pièce jointe (robuste) —
            If pieceJointe <> "" Then
                If LCase$(Left$(pieceJointe, 4)) = "http" Then
                    On Error Resume Next: .Attachments.Add pieceJointe: On Error GoTo 0
                ElseIf Dir(pieceJointe) <> "" Then
                    .Attachments.Add pieceJointe
                End If
            End If

            ' — Nettoyer les triples sauts de ligne avant la signature —
            Dim wdDoc As Object: Set wdDoc = .GetInspector.WordEditor
            With wdDoc.Content.Find
                .ClearFormatting
                .Text = "^p^p^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = 1
                .Execute Replace:=2
            End With

            .Save
        End With

        ' — Mise à jour de la date de dernière relance —
        ws.Cells(i, 11).Value = Date

SkipRow:
    Next i

    MsgBox cnt & " relance(s) préparée(s).", vbInformation
End Sub


