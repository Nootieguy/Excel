VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ark5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ============================================================================
' ARK5: AKTIVITETSOVERSIKT - WORKSHEET EVENT HANDLER
' ============================================================================
' FORMÅL:
'   Håndterer automatisk synkronisering mellom AKTIVITETSOVERSIKT og Planlegger
'   når brukeren endrer data direkte i oversikten.
'
' FUNKSJONALITET:
'   1. PERSON-SWITCHING: Endre person ? flytt aktivitet mellom personer
'   2. KODE-VALIDERING: Sjekk at aktivitetskoden finnes i AKTIVITETSTYPER
'   3. DATO-VALIDERING: Sjekk at startdato ikke er etter sluttdato
'   4. FORSINKELSE-HÅNDTERING: Respekter forsinkelse ved flytting
'
' REKURSJONSBESKYTTELSE:
'   - behandlerEndring flagg forhindrer evige loops
'   - Application.EnableEvents sjekk forhindrer meldinger under refresh
' ============================================================================

' ============================================================================
' SEKSJON 1: KONSTANTER OG GLOBALE VARIABLER
' ============================================================================
' Definerer tabellstruktur og beskyttelsesflagg

' Første rad med data i tabellen
Private Const TBL_START_ROW As Long = 4

' Kolonneindekser (hvilken kolonne inneholder hva)
Private Const COL_PERSON As Long = 1       ' A: Personnavn
Private Const COL_KODE As Long = 2         ' B: Aktivitetskode (TL, VA, osv)
Private Const COL_BESKR As Long = 3        ' C: Beskrivelse (auto-fylles)
Private Const COL_OPP_START As Long = 4    ' D: Opprinnelig startdato
Private Const COL_OPP_SLUTT As Long = 5    ' E: Opprinnelig sluttdato
Private Const COL_FORSINKET As Long = 6    ' F: Forsinkelse i dager (redigerbar)
Private Const COL_NY_SLUTT As Long = 7     ' G: Ny sluttdato (beregnet = E + F)
Private Const COL_KOMMENTAR As Long = 10   ' J: Kommentar (valgfri)

' Beskyttelsesflagg mot rekursjon
' Settes til True når vi behandler endring for å unngå nested calls
Private behandlerEndring As Boolean


' ============================================================================
' SEKSJON 2: HOVEDEVENT - WORKSHEET_CHANGE
' ============================================================================
' FORMÅL:
'   Fanger opp alle celleendringer i arket og delegerer til riktig handler
'
' FLYT:
'   1. Sjekk om vi allerede behandler (unngå rekursjon)
'   2. Disable events (unngå loops)
'   3. For hver endret celle: kall riktig handler basert på kolonne
'   4. Re-enable events
'
' KOLONNELOGIKK:
'   - Kolonne A (Person) ? HåndterPersonEndring
'   - Kolonne B (Kode) ? HåndterKodeEndring
'   - Kolonne D/E/F (Datoer/Forsinkelse) ? HåndterDatoEndring
' ============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Slutt
    
    ' STEG 1: Sjekk om vi allerede behandler (unngå rekursjon)
    If behandlerEndring Then Exit Sub
    
    ' STEG 2: Valider input
    If Target Is Nothing Then Exit Sub
    If Target.Row < TBL_START_ROW Then Exit Sub
    
    ' STEG 3: Sett beskyttelsesflagg
    behandlerEndring = True
    Application.EnableEvents = False
    
    ' STEG 4: Behandle hver endret celle
    Dim c As Range
    Dim antall As Long
    
    For Each c In Target.Cells
        antall = antall + 1
        
        ' Sikkerhet: maks 50 celler samtidig
        If antall > 50 Then
            If Application.EnableEvents Then
                MsgBox "For mange celler endret samtidig.", vbExclamation
            End If
            Exit For
        End If
        
        ' Kun behandle rader fra TBL_START_ROW og nedover
        If c.Row >= TBL_START_ROW Then
            ' Rute til riktig handler basert på hvilken kolonne som ble endret
            Select Case c.Column
                Case COL_PERSON
                    Call HåndterPersonEndring(c.Row)
                    
                Case COL_KODE
                    Call HåndterKodeEndring(c.Row)
                    
                Case COL_OPP_START, COL_OPP_SLUTT
                    Call HåndterDatoEndring(c.Row)
                    
                Case COL_FORSINKET
                    Call HåndterForsinkelseEndring(c.Row)
            End Select
        End If
    Next c
    
Slutt:
    ' STEG 5: Re-enable events og fjern beskyttelsesflagg
    Application.EnableEvents = True
    behandlerEndring = False
    
    ' Debug-logging ved feil
    If Err.Number <> 0 Then
        Debug.Print "Feil i Worksheet_Change: " & Err.Description
        Err.Clear
    End If
End Sub


' ============================================================================
' SEKSJON 3: PERSON-SWITCHING HANDLER
' ============================================================================
' FORMÅL:
'   Når brukeren endrer person-navnet i kolonne A, flyttes aktiviteten
'   automatisk fra gammel person til ny person i Planlegger-arket.
'
' FLYT:
'   1. Hent ny person, kode, datoer og forsinkelse fra raden
'   2. Beregn ny sluttdato (opprinnelig slutt + forsinkelse)
'   3. Finn gammel person ved å søke i Planlegger (3 metoder)
'   4. Sjekk om det faktisk er en endring
'   5. Kall Module3.FlyttAktivitetTilNyPerson for å utføre flyttingen
'
' SØKEMETODER (fallback-logikk):
'   Metode 1: Søk med ny slutt (hvis forsinket)
'   Metode 2: Søk med opprinnelig slutt (standard)
'   Metode 3: Bred søk (kun kode + startdato)
'
' FORSINKELSE-HÅNDTERING:
'   Hvis COL_FORSINKET > 0, brukes nySluttDato = sluttDato + forsinkelse
'   ved flytting til ny person.
' ============================================================================
Private Sub HåndterPersonEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler
    
    ' --- STEG 1: HENT DATA FRA RADEN ---
    Dim nyPerson As String      ' Ny person brukeren skrev inn
    Dim kode As String          ' Aktivitetskode (TL, VA, osv)
    Dim startDato As Date       ' Opprinnelig startdato
    Dim sluttDato As Date       ' Opprinnelig sluttdato
    Dim kommentar As String     ' Valgfri kommentar
    Dim forsinkelse As Long     ' Antall dager forsinket
    Dim nySluttDato As Date     ' Beregnet sluttdato (slutt + forsinkelse)
    Dim gammelPerson As String  ' Person som har aktiviteten nå
    Dim wsP As Worksheet        ' Planlegger-arket
    
    ' Hent person og kode
    nyPerson = Trim$(Me.Cells(rad, COL_PERSON).Value)
    kode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))
    
    ' Valider at vi har nødvendig data
    If Len(nyPerson) = 0 Or Len(kode) = 0 Then Exit Sub
    
    ' Hent datoer og forsinkelse (On Error Resume Next for å håndtere tomme celler)
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    forsinkelse = Me.Cells(rad, COL_FORSINKET).Value
    kommentar = Trim$(Me.Cells(rad, COL_KOMMENTAR).Value)
    On Error GoTo ErrorHandler
    
    ' --- STEG 2: BEREGN NY SLUTTDATO (MED FORSINKELSE) ---
    If forsinkelse > 0 Then
        nySluttDato = sluttDato + forsinkelse
    Else
        nySluttDato = sluttDato
    End If
    
    ' Valider datoer
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub
    
    ' --- STEG 3: FINN PLANLEGGER-ARKET ---
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets("Planlegger")
    On Error GoTo ErrorHandler
    
    If wsP Is Nothing Then
        If Application.EnableEvents Then
            MsgBox "Finner ikke Planlegger-arket!", vbCritical
        End If
        Exit Sub
    End If
    
    ' --- STEG 4: FINN GAMMEL PERSON (3 SØKEMETODER) ---
    ' Metode 1: Søk med ny slutt (hvis forsinket)
    If forsinkelse > 0 Then
        gammelPerson = FinnPersonForAktivitet(wsP, kode, startDato, nySluttDato)
    End If
    
    ' Metode 2: Søk med opprinnelig slutt (standard)
    If Len(gammelPerson) = 0 Then
        gammelPerson = FinnPersonForAktivitet(wsP, kode, startDato, sluttDato)
    End If
    
    ' Metode 3: Bred søk (kun kode og start - siste utvei)
    If Len(gammelPerson) = 0 Then
        gammelPerson = FinnPersonForAktivitetBred(wsP, kode, startDato)
    End If
    
    ' Hvis vi fortsatt ikke fant aktiviteten, vis feilmelding
    If Len(gammelPerson) = 0 Then
        If Application.EnableEvents Then
            MsgBox "Fant ikke aktiviteten i Planlegger!" & vbCrLf & vbCrLf & _
                   "Kode: " & kode & vbCrLf & _
                   "Start: " & Format(startDato, "dd.mm.yyyy") & vbCrLf & _
                   "Slutt: " & Format(sluttDato, "dd.mm.yyyy") & vbCrLf & vbCrLf & _
                   "Mulige årsaker:" & vbCrLf & _
                   "• Aktiviteten er slettet fra Planlegger" & vbCrLf & _
                   "• Datoene matcher ikke" & vbCrLf & _
                   "• Koden er endret", vbExclamation, "Kan ikke flytte"
        End If
        Exit Sub
    End If
    
    ' --- STEG 5: SJEKK OM DET ER EN FAKTISK ENDRING ---
    If StrComp(gammelPerson, nyPerson, vbTextCompare) = 0 Then Exit Sub
    
    ' --- STEG 6: FLYTT AKTIVITET TIL NY PERSON ---
    ' Kaller Module3.FlyttAktivitetTilNyPerson som utfører selve flyttingen
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!FlyttAktivitetTilNyPerson", _
                    gammelPerson, nyPerson, kode, _
                    startDato, nySluttDato, kommentar
    
    If Err.Number <> 0 Then
        If Application.EnableEvents Then
            MsgBox "Feil ved flytting av aktivitet:" & vbCrLf & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf & _
                   "Kontroller at Module3.FlyttAktivitetTilNyPerson finnes.", vbCritical
        End If
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    If Application.EnableEvents Then
        MsgBox "Feil ved person-endring:" & vbCrLf & Err.Description, vbCritical
    End If
    Debug.Print "FEIL i HåndterPersonEndring: " & Err.Description
End Sub


' ============================================================================
' SEKSJON 4: SØKEFUNKSJONER FOR AKTIVITETER I PLANLEGGER
' ============================================================================
' Disse funksjonene søker i Planlegger-arket for å finne hvilken person
' som har en gitt aktivitet.
' ============================================================================

' ----------------------------------------------------------------------------
' FUNKSJON: FinnPersonForAktivitet (EKSAKT MATCH)
' ----------------------------------------------------------------------------
' FORMÅL:
'   Finn personen som har en aktivitet med eksakt kode, startdato og sluttdato
'
' FLYT:
'   1. Hent Named Ranges (FirstDate, PersonHeader) for å finne struktur
'   2. Finn datokolonner for start og slutt
'   3. Søk gjennom alle personer (hovedrader med navn i kol A)
'   4. For hver person, sjekk hele personblokken (hovedrad + under-rader)
'   5. Hvis aktivitet funnet med riktig kode og spenn: returner personnavn
'
' SØKELOGIKK:
'   - Starter i startCol, sjekker om cellen har fet tekst med riktig kode
'   - Verifiserer at hele spennet (startCol til sluttCol) er samme aktivitet
'   - Sjekker både hovedrad og alle under-rader i personblokken
' ----------------------------------------------------------------------------
Private Function FinnPersonForAktivitet(wsP As Worksheet, _
                                       kode As String, _
                                       startDato As Date, _
                                       sluttDato As Date) As String
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long, blockEnd As Long
    Dim startCol As Long, sluttCol As Long
    Dim cel As Range
    Dim personNavn As String
    
    ' Hent struktur fra Named Ranges
    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0
    
    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then Exit Function
    
    ' Finn kolonne-indekser for start og slutt
    startCol = FinnDatoKol(wsP, startDato, datoRad, førsteDatoKol)
    sluttCol = FinnDatoKol(wsP, sluttDato, datoRad, førsteDatoKol)
    
    If startCol = 0 Or sluttCol = 0 Then Exit Function
    
    ' Søk gjennom alle personer
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    
    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)
        
        ' Kun personrader (har navn i kol A)
        If Len(personNavn) > 0 Then
            ' Finn hele personblokken (hovedrad + under-rader uten navn)
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop
            
            ' Søk i hele personblokken
            Dim blockRow As Long
            For blockRow = r To blockEnd
                Set cel = wsP.Cells(blockRow, startCol)
                
                ' Sjekk om dette er start på en aktivitet
                If cel.Font.Bold And Len(Trim$(cel.Value)) > 0 Then
                    ' Sjekk om koden matcher (case-insensitiv, starter med kode)
                    If InStr(1, Trim$(cel.Value), kode, vbTextCompare) = 1 Then
                        ' Verifiser at hele spennet er samme aktivitet
                        If VerifiserAktivitetSpenn(wsP, blockRow, startCol, sluttCol) Then
                            FinnPersonForAktivitet = personNavn
                            Exit Function
                        End If
                    End If
                End If
            Next blockRow
        End If
    Next r
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: FinnPersonForAktivitetBred (BRED SØKNING)
' ----------------------------------------------------------------------------
' FORMÅL:
'   Finn personen som har en aktivitet med gitt kode og startdato
'   (ignorerer sluttdato - brukes som fallback når eksakt match feiler)
'
' FORSKJELL FRA FinnPersonForAktivitet:
'   - Sjekker kun startkolonne, ikke sluttkolonne
'   - Verifiserer ikke hele spennet
'   - Brukes når forsinkelse gjør at sluttdato ikke matcher
' ----------------------------------------------------------------------------
Private Function FinnPersonForAktivitetBred(wsP As Worksheet, _
                                           kode As String, _
                                           startDato As Date) As String
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long, blockEnd As Long
    Dim startCol As Long
    Dim cel As Range
    Dim personNavn As String
    
    ' Hent struktur
    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0
    
    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then Exit Function
    
    ' Finn startkolonne
    startCol = FinnDatoKol(wsP, startDato, datoRad, førsteDatoKol)
    If startCol = 0 Then Exit Function
    
    ' Søk gjennom alle personer
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    
    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)
        If Len(personNavn) > 0 Then
            ' Finn hele personblokken
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop
            
            ' Søk i hele personblokken
            Dim blockRow As Long
            For blockRow = r To blockEnd
                Set cel = wsP.Cells(blockRow, startCol)
                
                ' Sjekk kun om koden matcher (ikke hele spennet)
                If cel.Font.Bold And Len(Trim$(cel.Value)) > 0 Then
                    If InStr(1, Trim$(cel.Value), kode, vbTextCompare) = 1 Then
                        FinnPersonForAktivitetBred = personNavn
                        Exit Function
                    End If
                End If
            Next blockRow
        End If
    Next r
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: VerifiserAktivitetSpenn
' ----------------------------------------------------------------------------
' FORMÅL:
'   Verifiser at alle celler fra startCol til sluttCol tilhører samme aktivitet
'
' LOGIKK:
'   - Hent fargen fra startcellen
'   - Sjekk at alle celler i spennet enten har fet tekst ELLER samme farge
'   - Hvis en celle ikke matcher ? ikke samme aktivitet
'
' BRUKES AV:
'   FinnPersonForAktivitet for å sikre eksakt match
' ----------------------------------------------------------------------------
Private Function VerifiserAktivitetSpenn(wsP As Worksheet, _
                                        rad As Long, _
                                        startCol As Long, _
                                        sluttCol As Long) As Boolean
    Dim c As Long
    Dim cel As Range
    Dim startFarge As Long
    Dim alleFet As Boolean
    
    ' Hent startfargen
    startFarge = wsP.Cells(rad, startCol).Interior.Color
    alleFet = True
    
    ' Sjekk alle celler i spennet
    For c = startCol To sluttCol
        Set cel = wsP.Cells(rad, c)
        
        ' Celle må enten ha fet tekst ELLER samme farge som start
        If Not cel.Font.Bold And cel.Interior.Color <> startFarge Then
            alleFet = False
            Exit For
        End If
    Next c
    
    VerifiserAktivitetSpenn = alleFet
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: FinnDatoKol
' ----------------------------------------------------------------------------
' FORMÅL:
'   Finn kolonneindeks for en gitt dato i datoraden
'
' FLYT:
'   Går gjennom alle kolonner fra førsteDatoKol til siste brukte kolonne
'   og sammenligner datoverdier
'
' RETURVERDI:
'   Kolonneindeks (Long) hvis funnet, 0 hvis ikke funnet
' ----------------------------------------------------------------------------
Private Function FinnDatoKol(wsP As Worksheet, _
                            søkDato As Date, _
                            datoRad As Long, _
                            førsteDatoKol As Long) As Long
    Dim lastCol As Long, c As Long
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
    
    For c = førsteDatoKol To lastCol
        If IsDate(wsP.Cells(datoRad, c).Value) Then
            ' CLng sammenligner bare datodelen, ignorerer tid
            If CLng(CDate(wsP.Cells(datoRad, c).Value)) = CLng(søkDato) Then
                FinnDatoKol = c
                Exit Function
            End If
        End If
    Next c
End Function


' ============================================================================
' SEKSJON 5: KODE-VALIDERING HANDLER
' ============================================================================
' FORMÅL:
'   Når brukeren endrer aktivitetskode i kolonne B, valideres koden mot
'   AKTIVITETSTYPER - OVERSIKT arket. Hvis koden er ugyldig, vises feilmelding
'   og cellen tømmes. Hvis koden er gyldig, fylles beskrivelsen automatisk.
'
' FLYT:
'   1. Hent ny kode fra cellen (UPPERCASE)
'   2. Finn AKTIVITETSTYPER - OVERSIKT arket
'   3. Slå opp koden via LookupAktivitet
'   4. Hvis ikke funnet: vis feilmelding (kun hvis EnableEvents)
'   5. Hvis funnet: fyll beskrivelse i kolonne C
'
' ENABLEEVENTS SJEKK:
'   Hvis Application.EnableEvents = False (under refresh), hoppes meldingsboks
'   og tømming over for å unngå evig loop.
' ============================================================================
Private Sub HåndterKodeEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler
    
    Dim nyKode As String
    Dim beskrivelse As String
    Dim farge As Long
    Dim wsTyp As Worksheet
    Dim funnet As Boolean
    
    ' STEG 1: Hent koden (UPPERCASE for konsistens)
    nyKode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))
    If Len(nyKode) = 0 Then Exit Sub
    
    ' STEG 2: Finn AKTIVITETSTYPER - OVERSIKT arket
    On Error Resume Next
    Set wsTyp = ThisWorkbook.Worksheets("AKTIVITETSTYPER - OVERSIKT")
    On Error GoTo ErrorHandler
    
    If wsTyp Is Nothing Then Exit Sub
    
    ' STEG 3: Slå opp koden
    On Error Resume Next
    funnet = Application.Run("LookupAktivitet", wsTyp, nyKode, beskrivelse, farge)
    On Error GoTo ErrorHandler
    
    ' STEG 4: Håndter resultat
    If Not funnet Then
        ' Koden finnes ikke - vis feilmelding kun hvis vi er i manuell modus
        If Application.EnableEvents Then
            MsgBox "Aktivitetskode '" & nyKode & "' finnes ikke i AKTIVITETSTYPER - OVERSIKT!" & vbCrLf & vbCrLf & _
                   "Legg til koden først, eller velg en eksisterende kode.", vbExclamation, "Ugyldig kode"
            Me.Cells(rad, COL_KODE).Value = ""
        End If
    Else
        ' Koden finnes - fyll beskrivelse automatisk
        Me.Cells(rad, COL_BESKR).Value = beskrivelse
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Feil i HåndterKodeEndring: " & Err.Description
    Application.EnableEvents = True
End Sub


' ============================================================================
' SEKSJON 6A: FORSINKELSE-ENDRING HANDLER
' ============================================================================
' FORMÅL:
'   Når brukeren endrer forsinkelse i kolonne F, må aktiviteten i Planlegger
'   oppdateres med ny lengde.
'
' PROBLEM DET LØSER:
'   - Uten denne: aktiviteten i Planlegger forblir med gammel lengde
'   - Gamle celler blir ikke ryddet ? røde skraveringer
'   - Status oppdateres ikke
'
' FLYT:
'   1. Hent person, kode, datoer og forsinkelse
'   2. Beregn ny sluttdato (slutt + forsinkelse)
'   3. Kall SynkroniserEnkeltAktivitet for å re-tegne aktiviteten
'
' SYNKRONISER VS FLYTT:
'   - SynkroniserEnkeltAktivitet: Rydder gammel + tegner ny for samme person
'   - FlyttAktivitetTilNyPerson: Flytter mellom forskjellige personer
' ============================================================================
Private Sub HåndterForsinkelseEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler
    
    ' --- STEG 1: HENT DATA ---
    Dim person As String
    Dim kode As String
    Dim startDato As Date
    Dim sluttDato As Date
    Dim forsinkelse As Long
    Dim nySluttDato As Date
    Dim kommentar As String
    
    person = Trim$(Me.Cells(rad, COL_PERSON).Value)
    kode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))
    
    If Len(person) = 0 Or Len(kode) = 0 Then Exit Sub
    
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    forsinkelse = Me.Cells(rad, COL_FORSINKET).Value
    kommentar = Trim$(Me.Cells(rad, COL_KOMMENTAR).Value)
    On Error GoTo ErrorHandler
    
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub
    
    ' --- STEG 2A: VERIFISER AT AKTIVITETEN FINNES I PLANLEGGER FØRST ---
    ' Dette forhindrer unødvendige feilmeldinger
    Dim wsP As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets("Planlegger")
    On Error GoTo ErrorHandler
    
    If wsP Is Nothing Then Exit Sub
    
    ' Prøv å finne aktiviteten først (stille sjekk)
    Dim testPerson As String
    testPerson = FinnPersonForAktivitetBred(wsP, kode, startDato)
    
    ' Hvis aktiviteten ikke finnes, hopp over stille
    If Len(testPerson) = 0 Then
        Debug.Print "HåndterForsinkelseEndring: Aktivitet " & kode & " ikke funnet i Planlegger - hopper over"
        Exit Sub
    End If
    
    ' Hvis person ikke matcher, hopp over stille
    If StrComp(testPerson, person, vbTextCompare) <> 0 Then
        Debug.Print "HåndterForsinkelseEndring: Person mismatch - forventet " & person & ", fant " & testPerson & " - hopper over"
        Exit Sub
    End If
    
    ' --- STEG 2B: BEREGN NY SLUTTDATO ---
    If forsinkelse > 0 Then
        nySluttDato = sluttDato + forsinkelse
    Else
        nySluttDato = sluttDato
    End If
    
    ' --- STEG 3: SYNKRONISER AKTIVITETEN I PLANLEGGER ---
    ' SynkroniserEnkeltAktivitet rydder gammel aktivitet og tegner ny
    ' VIKTIG: Midlertidig enable events for å la SynkroniserEnkeltAktivitet vise meldinger
    '         bare hvis vi er i manuell modus (EnableEvents var True til å begynne med)
    Dim opprinneligEventsStatus As Boolean
    opprinneligEventsStatus = Application.EnableEvents
    
    ' Re-enable events midlertidig hvis vi var i manuell modus
    ' (dette lar SynkroniserEnkeltAktivitet vise viktige feilmeldinger)
    If opprinneligEventsStatus Then
        Application.EnableEvents = True
    End If
    
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!SynkroniserEnkeltAktivitet", _
                    person, kode, startDato, nySluttDato, kommentar
    
    If Err.Number <> 0 Then
        If opprinneligEventsStatus Then
            MsgBox "Feil ved oppdatering av forsinkelse:" & vbCrLf & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf & _
                   "Kontroller at Module3.SynkroniserEnkeltAktivitet finnes.", vbCritical
        End If
        Err.Clear
    End If
    
    ' Restore original events status
    Application.EnableEvents = opprinneligEventsStatus
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    If Application.EnableEvents Then
        MsgBox "Feil ved forsinkelse-endring:" & vbCrLf & Err.Description, vbCritical
    End If
    Debug.Print "FEIL i HåndterForsinkelseEndring: " & Err.Description
End Sub


' ============================================================================
' SEKSJON 6B: DATO-VALIDERING HANDLER
' ============================================================================
' FORMÅL:
'   Når brukeren endrer startdato, sluttdato eller forsinkelse, valideres
'   at startdato ikke er etter sluttdato.
'
' FLYT:
'   1. Hent startdato og sluttdato
'   2. Sjekk om begge er gyldige datoer
'   3. Hvis startdato > sluttdato: vis feilmelding og reset startdato
'
' ENABLEEVENTS SJEKK:
'   Hvis Application.EnableEvents = False, hoppes meldingsboks og reset
'   over for å unngå problemer under refresh.
' ============================================================================
Private Sub HåndterDatoEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler
    
    Dim startDato As Date
    Dim sluttDato As Date
    
    ' STEG 1: Hent datoer
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    On Error GoTo ErrorHandler
    
    ' STEG 2: Valider at begge er datoer
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub
    
    ' STEG 3: Valider at startdato ikke er etter sluttdato
    If startDato > sluttDato Then
        ' Vis feilmelding kun hvis vi er i manuell modus
        If Application.EnableEvents Then
            MsgBox "Startdato kan ikke være etter sluttdato!", vbExclamation, "Ugyldig dato"
            Me.Cells(rad, COL_OPP_START).Value = sluttDato
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Feil i HåndterDatoEndring: " & Err.Description
    Application.EnableEvents = True
End Sub


' ============================================================================
' SEKSJON 7: UTILITY-FUNKSJONER
' ============================================================================
' Hjelpefunksjoner for debugging og feilsøking
' ============================================================================

' ----------------------------------------------------------------------------
' FUNKSJON: ResetEventHandler
' ----------------------------------------------------------------------------
' FORMÅL:
'   Manuell reset av event handler ved problemer
'
' BRUK:
'   Kjør denne i Immediate Window (Ctrl+G) hvis du opplever problemer:
'   Call Ark5.ResetEventHandler
'
' EFFEKT:
'   - Nullstiller behandlerEndring flagg
'   - Re-enabler events
'   - Logger til Debug window
' ----------------------------------------------------------------------------
Public Sub ResetEventHandler()
    behandlerEndring = False
    Application.EnableEvents = True
    Debug.Print "Event handler reset - Ark5 v3.1 (med forsinkelse-fix)"
End Sub


