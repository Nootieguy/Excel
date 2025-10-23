VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Private Sub Worksheet_Activate()
    modDagensOutline.MarkerDagensDato_Bla
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrHandler

    Debug.Print ">>> Worksheet_SelectionChange trigget! Target = " & Target.Address & ", CellCount = " & Target.Cells.Count

    ' Håndter kant-klikk resize (to-stegs: først kant, så målcelle)
    If Target.Cells.Count = 1 Then
        Debug.Print ">>> Kaller Module5.HaandterCelleKlikk..."
        Call Module5.HaandterCelleKlikk(Me, Target)
        Debug.Print ">>> Tilbake fra Module5.HaandterCelleKlikk"
    Else
        Debug.Print ">>> Target.Cells.Count = " & Target.Cells.Count & " (ikke 1, skipper)"
    End If

    Exit Sub

ErrHandler:
    Debug.Print ">>> FEIL i Worksheet_SelectionChange: " & Err.Description
End Sub

