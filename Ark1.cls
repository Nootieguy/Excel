VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Private Sub Worksheet_Activate()
    modDagensOutline.MarkerDagensDato_Bla
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrHandler

    ' Dobbeltklikk for å resize aktivitet
    Debug.Print "BeforeDoubleClick: Target = " & Target.Address

    ' Kall resize-funksjonen
    Call Module5.HaandterDobbeltklikk(Me, Target)

    ' Avbryt default-oppførsel (gå inn i edit mode)
    Cancel = True

    Exit Sub

ErrHandler:
    Debug.Print "FEIL i Worksheet_BeforeDoubleClick: " & Err.Description
End Sub

