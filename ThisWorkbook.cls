VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call MasquerLignesRefusees

    ' Ne mettre � jour la liste des contacts que si elle n'existe pas
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("ListeContacts")
    On Error GoTo 0

    If ws Is Nothing Then
        Call MettreAJourListeContacts
    End If
End Sub

