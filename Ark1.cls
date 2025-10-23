VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Private Sub Worksheet_Activate()
    modDagensOutline.MarkerDagensDato_Bla
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrHandler

    ' Sjekk om Shift er holdt nede og brukeren klikket på én celle
    If Target.Cells.Count = 1 Then
        ' Windows API: GetKeyState returnerer negativ verdi hvis tasten er nede
        ' VK_SHIFT = 16
        Dim shiftPressed As Boolean
        shiftPressed = (Module5.GetAsyncKeyState(16) And &H8000) <> 0

        Debug.Print "SelectionChange: Shift pressed = " & shiftPressed & ", Target = " & Target.Address

        If shiftPressed Then
            Debug.Print "Kaller Module5.HaandterShiftKlikk..."
            ' Kall resize-funksjonen
            Call Module5.HaandterShiftKlikk(Me, Target)
        End If
    End If
    Exit Sub

ErrHandler:
    Debug.Print "FEIL i Worksheet_SelectionChange: " & Err.Description
End Sub

