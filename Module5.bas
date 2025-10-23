Option Explicit

' ==============================================================================
' MODULE 5 - RESIZE-FUNKSJONALITET FOR MERGED AKTIVITETER
' ==============================================================================
' Håndterer kant-klikk drag resize av merged aktiviteter i Planlegger-arket
' Automatisk overlapp-håndtering og rutenett-restore
'
' BRUK:
'   1. Klikk på KANTEN av en merged aktivitet (første eller siste kolonne)
'   2. Klikk på målcellen hvor du vil kanten skal flyttes
'
' FUNKSJONER:
'   - HaandterCelleKlikk: Hovedfunksjon for to-stegs resize
'   - ErKantAvMergedAktivitet: Sjekker om klikk er på kant
'   - VisResizeModus: Visuell feedback for resize-modus
'   - HaandterDobbeltklikk: Legacy (kan fjernes senere)
' ==============================================================================

' ===== RESIZE STATE =====
Private resizeModus As Boolean
Private resizeRad As Long
Private resizeStartKol As Long
Private resizeSluttKol As Long
Private resizeVenstre As Boolean ' True = resize venstre kant, False = høyre kant
Private resizeTekst As String
Private resizeFarge As Long

' ----------------------------------------------------------------------------
' FUNKSJON: HaandterCelleKlikk (To-stegs resize)
' ----------------------------------------------------------------------------
' STEG 1: Klikk på kant av aktivitet → Aktiverer resize-modus
' STEG 2: Klikk på målcelle → Utfører resize
' ----------------------------------------------------------------------------
Public Sub HaandterCelleKlikk(wsP As Worksheet, Target As Range)
    On Error GoTo ErrHandler

    Debug.Print "=== HaandterCelleKlikk: " & Target.Address & " ==="

    If Not resizeModus Then
        ' STEG 1: Sjekk om brukeren klikket på en kant
        Dim kantInfo As Object
        Set kantInfo = ErKantAvMergedAktivitet(wsP, Target)

        If Not kantInfo Is Nothing Then
            ' Aktivér resize-modus
            resizeModus = True
            resizeRad = kantInfo("Rad")
            resizeStartKol = kantInfo("StartKol")
            resizeSluttKol = kantInfo("SluttKol")
            resizeVenstre = kantInfo("VenstreKant")
            resizeTekst = kantInfo("Tekst")
            resizeFarge = kantInfo("Farge")

            ' Visuell feedback
            Call VisResizeModus(wsP, True)

            Dim kantNavn As String
            kantNavn = IIf(resizeVenstre, "venstre", "høyre")
            Application.StatusBar = "RESIZE-MODUS: Klikk på målcellen for å flytte " & kantNavn & " kant"
            Debug.Print "Resize-modus aktivert: " & kantNavn & " kant av aktivitet i rad " & resizeRad
        End If
    Else
        ' STEG 2: Brukeren klikket på målcelle - utfør resize
        Dim malKol As Long
        malKol = Target.Column

        Debug.Print "Målcelle: kolonne " & malKol

        ' Beregn nye dimensjoner
        Dim nyStartKol As Long, nySluttKol As Long
        If resizeVenstre Then
            ' Resize venstre kant
            nyStartKol = malKol
            nySluttKol = resizeSluttKol
        Else
            ' Resize høyre kant
            nyStartKol = resizeStartKol
            nySluttKol = malKol
        End If

        ' Sjekk at det er en gyldig endring
        If nyStartKol < nySluttKol And nyStartKol <> resizeStartKol Or nySluttKol <> resizeSluttKol Then
            ' Håndter overlapp
            Call HaandterOverlappVedResize(wsP, resizeRad, nyStartKol, nySluttKol, resizeStartKol, resizeSluttKol)

            ' Utfør resize
            Call ResizeMergedAktivitet(wsP, resizeRad, resizeStartKol, resizeSluttKol, nyStartKol, nySluttKol, resizeTekst, resizeFarge)

            Application.StatusBar = "Aktivitet resized til " & (nySluttKol - nyStartKol + 1) & " dager"
            Debug.Print "Resize fullført!"
        Else
            Application.StatusBar = "Ugyldig resize - avbrutt"
            Debug.Print "Ugyldig resize (nyStart=" & nyStartKol & ", nyS lutt=" & nySluttKol & ")"
        End If

        ' Deaktivér resize-modus
        Call VisResizeModus(wsP, False)
        resizeModus = False
    End If

    Exit Sub

ErrHandler:
    Debug.Print "FEIL i HaandterCelleKlikk: " & Err.Description
    resizeModus = False
    Call VisResizeModus(wsP, False)
    MsgBox "Feil ved resize: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: ErKantAvMergedAktivitet
' ----------------------------------------------------------------------------
' Sjekker om en celle er kanten (første eller siste kolonne) av en merged aktivitet
' Returnerer Dictionary med info hvis det er en kant, ellers Nothing
' ----------------------------------------------------------------------------
Private Function ErKantAvMergedAktivitet(wsP As Worksheet, Target As Range) As Object
    On Error Resume Next

    Debug.Print "  ErKantAvMergedAktivitet: Sjekker " & Target.Address

    Dim cel As Range
    Set cel = Target.Cells(1, 1) ' Første celle hvis multi-select

    ' Sjekk om cellen er merged
    If Not cel.MergeCells Then
        Debug.Print "    → Ikke merged"
        Set ErKantAvMergedAktivitet = Nothing
        Exit Function
    End If

    Debug.Print "    → Er merged"

    Dim ma As Range
    Set ma = cel.MergeArea

    ' Debug merged area info
    Debug.Print "    → MergeArea: " & ma.Address & ", Farge: " & ma.Interior.Color & ", Bold: " & ma.Font.Bold & ", Verdi: [" & ma.Value & "]"

    ' Sjekk om det er en aktivitet (har farge og fet tekst)
    If ma.Interior.Color = RGB(255, 255, 255) Or _
       ma.Interior.ColorIndex = xlColorIndexNone Or _
       Not ma.Font.Bold Or _
       Len(Trim$(ma.Value)) = 0 Then
        Debug.Print "    → Ikke en aktivitet (hvit/ingen farge ELLER ikke bold ELLER tom)"
        Set ErKantAvMergedAktivitet = Nothing
        Exit Function
    End If

    Debug.Print "    → Er en aktivitet!"

    ' Sjekk om dette er første eller siste kolonne i merged area
    Dim startKol As Long, sluttKol As Long, klikketKol As Long
    startKol = ma.Column
    sluttKol = ma.Column + ma.Columns.Count - 1
    klikketKol = cel.Column

    Debug.Print "    → StartKol: " & startKol & ", SluttKol: " & sluttKol & ", KlikketKol: " & klikketKol

    Dim erVenstre As Boolean, erHoyre As Boolean
    erVenstre = (klikketKol = startKol)
    erHoyre = (klikketKol = sluttKol)

    Debug.Print "    → ErVenstre: " & erVenstre & ", ErHoyre: " & erHoyre

    If Not erVenstre And Not erHoyre Then
        ' Ikke en kant
        Debug.Print "    → IKKE EN KANT (midten av aktivitet)"
        Set ErKantAvMergedAktivitet = Nothing
        Exit Function
    End If

    ' Det er en kant! Returner info
    Debug.Print "    → ✓ DET ER EN KANT!"
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info("Rad") = ma.Row
    info("StartKol") = startKol
    info("SluttKol") = sluttKol
    info("VenstreKant") = erVenstre
    info("Tekst") = CStr(ma.Value)
    info("Farge") = ma.Interior.Color

    Set ErKantAvMergedAktivitet = info
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: VisResizeModus
' ----------------------------------------------------------------------------
' Visuell feedback når resize-modus er aktiv
' ----------------------------------------------------------------------------
Private Sub VisResizeModus(wsP As Worksheet, aktiv As Boolean)
    On Error Resume Next

    Debug.Print "  VisResizeModus: aktiv=" & aktiv

    If aktiv Then
        ' Highlight kanten som skal resizes
        Dim kantKol As Long
        If resizeVenstre Then
            kantKol = resizeStartKol
        Else
            kantKol = resizeSluttKol
        End If

        Debug.Print "    → Setter rød border på rad " & resizeRad & ", kolonne " & kantKol

        With wsP.Cells(resizeRad, kantKol)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeLeft).Color = RGB(255, 0, 0) ' Rød
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThick
            .Borders(xlEdgeRight).Color = RGB(255, 0, 0)
        End With

        Debug.Print "    → Rød border satt!"
    Else
        ' Fjern highlight (reset til merged cell sin original border)
        Debug.Print "    → Deaktiverer resize-modus"
        Application.StatusBar = False
    End If
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: HaandterDobbeltklikk (LEGACY - KAN FJERNES)
' ----------------------------------------------------------------------------
' Håndterer dobbeltklikk for å resize merged aktiviteter
' Brukeren dobbeltklikker på målcellen (hvor aktiviteten skal utvides til)
' ----------------------------------------------------------------------------
Public Sub HaandterDobbeltklikk(wsP As Worksheet, Target As Range)
    On Error GoTo ErrHandler

    Debug.Print "=== HaandterDobbeltklikk START ==="

    Dim rad As Long, malKol As Long
    rad = Target.Row
    malKol = Target.Column

    Debug.Print "Rad: " & rad & ", Kolonne: " & malKol

    ' Finn nærmeste merged aktivitet i samme rad
    Dim aktivitetInfo As Object
    Set aktivitetInfo = FinnNaermesteMergedAktivitet(wsP, rad, malKol)

    If aktivitetInfo Is Nothing Then
        ' Ingen aktivitet funnet i denne raden
        Debug.Print "Ingen aktivitet funnet i rad " & rad
        Exit Sub
    End If

    Debug.Print "Aktivitet funnet!"

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
    Debug.Print "FEIL i HaandterDobbeltklikk: " & Err.Description
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
