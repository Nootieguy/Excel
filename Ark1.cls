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

    ' Håndter kant-klikk resize (fungerer også for merged cells!)
    Debug.Print ">>> Kaller Module5.HaandterCelleKlikk..."
    Call Module5.HaandterCelleKlikk(Me, Target)
    Debug.Print ">>> Tilbake fra Module5.HaandterCelleKlikk"

    Exit Sub

ErrHandler:
    Debug.Print ">>> FEIL i Worksheet_SelectionChange: " & Err.Description
End Sub

