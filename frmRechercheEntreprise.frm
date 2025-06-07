VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRechercheEntreprise 
   Caption         =   "Sélection d'une entreprise"
   ClientHeight    =   1645
   ClientLeft      =   -343
   ClientTop       =   -1771
   ClientWidth     =   2016
   OleObjectBlob   =   "frmRechercheEntreprise.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRechercheEntreprise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSelectionner_Click()
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner une entreprise.", vbExclamation
        Exit Sub
    End If

    ActiveCell.Value = Me.ListBox1.Value
    Feuil1.TraiterEntreprise ActiveCell
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 100
    Me.Left = Application.Left + 100
    Me.Width = 360
    Me.Height = 220

    Dim ws As Worksheet
    Dim cell As range
    Dim entrepriseRecherchee As String
    Dim entreprisesTrouvees As Collection
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("ListeContacts")
    entrepriseRecherchee = LCase(Trim(ActiveCell.Value))
    Set entreprisesTrouvees = New Collection

    On Error Resume Next
    For Each cell In ws.range("A2:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).row)
        If LCase(cell.Value) Like "*" & entrepriseRecherchee & "*" Then
            entreprisesTrouvees.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0

    If entreprisesTrouvees.Count = 0 Then
        MsgBox "Aucune entreprise trouvée pour """ & entrepriseRecherchee & """", vbExclamation
        Unload Me
        Exit Sub
    End If

    For i = 1 To entreprisesTrouvees.Count
        Me.ListBox1.AddItem entreprisesTrouvees(i)
    Next i

    ' Ne rien faire d'autre – pas de repositionnement du bouton

    DoEvents
    Me.Repaint
End Sub

