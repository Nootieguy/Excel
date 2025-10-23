Attribute VB_Name = "Module5"
Option Explicit

' ==============================================================================
' MODULE 5 - RESIZE-FUNKSJONALITET FOR MERGED AKTIVITETER
' ==============================================================================
' Håndterer Ctrl+klikk resize av merged aktiviteter i Planlegger-arket
' Automatisk overlapp-håndtering og rutenett-restore
'
' BRUK:
'   Hold Ctrl og klikk på målcellen for å resize nærmeste aktivitet
'
' FUNKSJONER:
'   - HaandterCtrlKlikk: Hovedfunksjon som håndterer Ctrl+klikk
'   - FinnNaermesteMergedAktivitet: Finner aktivitet å resize
'   - ResizeMergedAktivitet: Utfører resize med rutenett-restore
'   - HaandterOverlappVedResize: Håndterer kolliderende aktiviteter
' ==============================================================================

' ===== WINDOWS API FOR CTRL-KLIKK DETEKSJON =====
#If VBA7 Then
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' ----------------------------------------------------------------------------
' FUNKSJON: HaandterCtrlKlikk
' ----------------------------------------------------------------------------
' Håndterer Ctrl+klikk for å resize merged aktiviteter
' Brukeren holder Ctrl og klikker på målcellen (hvor aktiviteten skal utvides til)
' ----------------------------------------------------------------------------
Public Sub HaandterCtrlKlikk(wsP As Worksheet, Target As Range)
    On Error GoTo ErrHandler

    Dim rad As Long, malKol As Long
    rad = Target.Row
    malKol = Target.Column

    ' Finn nærmeste merged aktivitet i samme rad
    Dim aktivitetInfo As Object
    Set aktivitetInfo = FinnNaermesteMergedAktivitet(wsP, rad, malKol)

    If aktivitetInfo Is Nothing Then
        ' Ingen aktivitet funnet i denne raden
        Exit Sub
    End If

    ' Hent info om aktiviteten
    Dim startKol As Long, sluttKol As Long
    Dim aktivitetTekst As String, aktivitetFarge As Long
    startKol = aktivitetInfo("StartKol")
    sluttKol = aktivitetInfo("SluttKol")
    aktivitetTekst = aktivitetInfo("Tekst")
    aktivitetFarge = aktivitetInfo("Farge")

    ' Bestem hvilken kant som skal flyttes
    Dim nyStartKol As Long, nySluttKol As Long

    If malKol < startKol Then
        ' Utvid til venstre
        nyStartKol = malKol
        nySluttKol = sluttKol
    ElseIf malKol > sluttKol Then
        ' Utvid til høyre
        nyStartKol = startKol
        nySluttKol = malKol
    ElseIf malKol >= startKol And malKol <= sluttKol Then
        ' Klikket inne i aktiviteten - bestem nærmeste kant
        Dim avstandVenstre As Long, avstandHoyre As Long
        avstandVenstre = malKol - startKol
        avstandHoyre = sluttKol - malKol

        If avstandVenstre < avstandHoyre Then
            ' Flytt venstre kant
            nyStartKol = malKol
            nySluttKol = sluttKol
        Else
            ' Flytt høyre kant
            nyStartKol = startKol
            nySluttKol = malKol
        End If
    End If

    ' Sjekk at vi faktisk endrer størrelsen
    If nyStartKol = startKol And nySluttKol = sluttKol Then
        Exit Sub
    End If

    ' Håndter overlapp før resize
    Call HaandterOverlappVedResize(wsP, rad, nyStartKol, nySluttKol, startKol, sluttKol)

    ' Utfør resize
    Call ResizeMergedAktivitet(wsP, rad, startKol, sluttKol, nyStartKol, nySluttKol, aktivitetTekst, aktivitetFarge)

    ' Gi feedback
    Dim msg As String
    If nySluttKol - nyStartKol > sluttKol - startKol Then
        msg = "Aktiviteten utvidet"
    Else
        msg = "Aktiviteten redusert"
    End If
    Application.StatusBar = msg & " til " & (nySluttKol - nyStartKol + 1) & " dager"

    Exit Sub

ErrHandler:
    MsgBox "Feil ved resize: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: FinnNaermesteMergedAktivitet
' ----------------------------------------------------------------------------
' Finner nærmeste merged aktivitet i en gitt rad
' Returnerer Dictionary med StartKol, SluttKol, Tekst, Farge
' ----------------------------------------------------------------------------
Private Function FinnNaermesteMergedAktivitet(wsP As Worksheet, rad As Long, _
                                              malKol As Long) As Object
    On Error Resume Next

    Dim lastCol As Long
    lastCol = wsP.Cells(rad, wsP.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    Dim naermeste As Object
    Dim minAvstand As Long
    minAvstand = 9999999

    ' Skann alle celler i raden for merged aktiviteter
    For c = 1 To lastCol
        Dim cel As Range
        Set cel = wsP.Cells(rad, c)

        If cel.MergeCells Then
            Dim ma As Range
            Set ma = cel.MergeArea

            ' Sjekk om dette er en aktivitet (har farge og tekst)
            If ma.Interior.Color <> RGB(255, 255, 255) And _
               ma.Interior.ColorIndex <> xlColorIndexNone And _
               Len(Trim$(ma.Value)) > 0 Then

                Dim aktStartKol As Long, aktSluttKol As Long
                aktStartKol = ma.Column
                aktSluttKol = ma.Column + ma.Columns.Count - 1

                ' Beregn avstand til målkolonnen
                Dim avstand As Long
                If malKol < aktStartKol Then
                    avstand = aktStartKol - malKol
                ElseIf malKol > aktSluttKol Then
                    avstand = malKol - aktSluttKol
                Else
                    avstand = 0 ' Inne i aktiviteten
                End If

                ' Sjekk om dette er nærmeste aktivitet
                If avstand < minAvstand Then
                    minAvstand = avstand

                    Set naermeste = CreateObject("Scripting.Dictionary")
                    naermeste("StartKol") = aktStartKol
                    naermeste("SluttKol") = aktSluttKol
                    naermeste("Tekst") = CStr(ma.Value)
                    naermeste("Farge") = ma.Interior.Color
                End If

                ' Hopp over resten av merged area
                c = aktSluttKol
            End If
        End If
    Next c

    Set FinnNaermesteMergedAktivitet = naermeste
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: ResizeMergedAktivitet
' ----------------------------------------------------------------------------
' Resizer en merged aktivitet fra gammelt område til nytt område
' ----------------------------------------------------------------------------
Private Sub ResizeMergedAktivitet(wsP As Worksheet, rad As Long, _
                                  gammelStartKol As Long, gammelSluttKol As Long, _
                                  nyStartKol As Long, nySluttKol As Long, _
                                  tekst As String, farge As Long)
    On Error GoTo ErrHandler

    ' STEG 1: Slett gammelt merged område
    Call Module3.SlettMergedAktivitet(wsP, rad, gammelStartKol)

    ' STEG 2: Restore rutenett i områder som ikke lenger dekkes
    Dim c As Long
    If nyStartKol > gammelStartKol Then
        ' Frigjort område til venstre
        For c = gammelStartKol To nyStartKol - 1
            Call RestoreRutenett(wsP.Cells(rad, c))
        Next c
    End If

    If nySluttKol < gammelSluttKol Then
        ' Frigjort område til høyre
        For c = nySluttKol + 1 To gammelSluttKol
            Call RestoreRutenett(wsP.Cells(rad, c))
        Next c
    End If

    ' STEG 3: Lag ny merged aktivitet
    Call Module3.LagMergedAktivitet(wsP, rad, nyStartKol, nySluttKol, farge, tekst)

    Exit Sub

ErrHandler:
    MsgBox "Feil ved resize av aktivitet: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: HaandterOverlappVedResize
' ----------------------------------------------------------------------------
' Håndterer overlappende aktiviteter når en aktivitet resizes
' Reduserer eller fjerner aktiviteter som blir dekket
' ----------------------------------------------------------------------------
Private Sub HaandterOverlappVedResize(wsP As Worksheet, rad As Long, _
                                      nyStartKol As Long, nySluttKol As Long, _
                                      eksisterendeStartKol As Long, eksisterendeSluttKol As Long)
    On Error Resume Next

    Dim c As Long
    Dim behandledeAktiviteter As Object
    Set behandledeAktiviteter = CreateObject("Scripting.Dictionary")

    ' Skann området som vil bli dekket (ekskluder den eksisterende aktiviteten)
    For c = nyStartKol To nySluttKol
        ' Hopp over den eksisterende aktivitetens område
        If c >= eksisterendeStartKol And c <= eksisterendeSluttKol Then
            GoTo NesteCelle
        End If

        Dim cel As Range
        Set cel = wsP.Cells(rad, c)

        If cel.MergeCells Then
            Dim ma As Range
            Set ma = cel.MergeArea

            Dim aktStartKol As Long, aktSluttKol As Long
            aktStartKol = ma.Column
            aktSluttKol = ma.Column + ma.Columns.Count - 1

            ' Unngå å behandle samme aktivitet flere ganger
            Dim aktKey As String
            aktKey = CStr(aktStartKol) & "|" & CStr(aktSluttKol)

            If Not behandledeAktiviteter.exists(aktKey) Then
                behandledeAktiviteter.Add aktKey, True

                ' Sjekk om aktiviteten er inne i resize-området
                If aktStartKol >= nyStartKol And aktSluttKol <= nySluttKol Then
                    ' HELT DEKKET - Fjern aktiviteten
                    Call Module3.SlettMergedAktivitet(wsP, rad, aktStartKol)

                ElseIf aktStartKol < nyStartKol And aktSluttKol > nySluttKol Then
                    ' RESIZE-OMRÅDET ER INNE I AKTIVITETEN
                    ' Dette skal ikke skje hvis vi resizer en aktivitet som overlapper en annen
                    ' Men vi håndterer det likevel: Behold venstre del
                    Dim aktivitetTekst As String, aktivitetFarge As Long
                    aktivitetTekst = CStr(ma.Value)
                    aktivitetFarge = ma.Interior.Color

                    Call ResizeMergedAktivitet(wsP, rad, aktStartKol, aktSluttKol, _
                                              aktStartKol, nyStartKol - 1, _
                                              aktivitetTekst, aktivitetFarge)

                ElseIf aktStartKol < nyStartKol And aktSluttKol >= nyStartKol And aktSluttKol <= nySluttKol Then
                    ' OVERLAPP PÅ HØYRE SIDE - Reduser til venstre del
                    Dim akt2Tekst As String, akt2Farge As Long
                    akt2Tekst = CStr(ma.Value)
                    akt2Farge = ma.Interior.Color

                    Call ResizeMergedAktivitet(wsP, rad, aktStartKol, aktSluttKol, _
                                              aktStartKol, nyStartKol - 1, _
                                              akt2Tekst, akt2Farge)

                ElseIf aktStartKol >= nyStartKol And aktStartKol <= nySluttKol And aktSluttKol > nySluttKol Then
                    ' OVERLAPP PÅ VENSTRE SIDE - Reduser til høyre del
                    Dim akt3Tekst As String, akt3Farge As Long
                    akt3Tekst = CStr(ma.Value)
                    akt3Farge = ma.Interior.Color

                    Call ResizeMergedAktivitet(wsP, rad, aktStartKol, aktSluttKol, _
                                              nySluttKol + 1, aktSluttKol, _
                                              akt3Tekst, akt3Farge)
                End If

                ' Hopp over resten av denne aktiviteten
                c = aktSluttKol
            End If
        End If

NesteCelle:
    Next c
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: RestoreRutenett
' ----------------------------------------------------------------------------
' Restore rutenett i en celle som frigjøres
' ----------------------------------------------------------------------------
Private Sub RestoreRutenett(cel As Range)
    cel.ClearComments
    cel.ClearContents
    cel.Interior.ColorIndex = xlColorIndexNone
    cel.Font.Bold = False
    cel.Font.Color = RGB(0, 0, 0)
    cel.HorizontalAlignment = xlCenter
    cel.VerticalAlignment = xlCenter

    ' Restore rutenett
    With cel.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
