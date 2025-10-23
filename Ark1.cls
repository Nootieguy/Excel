VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Private Sub Worksheet_Activate()
    modDagensOutline.MarkerDagensDato_Bla
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrHandler

    ' Håndter kant-klikk resize (to-stegs: først kant, så målcelle)
    If Target.Cells.Count = 1 Then
        Call Module5.HaandterCelleKlikk(Me, Target)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "FEIL i Worksheet_SelectionChange: " & Err.Description
End Sub

