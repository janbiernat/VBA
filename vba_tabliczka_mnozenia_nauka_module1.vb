Dim PytA As Integer
Dim PytB As Integer
Dim OdpD As Integer
Dim OdpZ As Integer

Sub PytanieLosuj(PytanieTylko As Boolean)
'PytanieLosuj - Procedura losuje pytanie (np. 3 * 4).
'
    If (PytanieTylko = False) Then
        Randomize                       '1.
        PytA = 0
        PytA = Round((10 * Rnd), 0)  '2.
        PytB = 0
        PytB = Round((10 * Rnd), 0)
    End If
    Cells(5, 2).Value = CStr(PytA) + " x " + CStr(PytB) + " = "
    '
    'Legenda:
    '1. Uruchomienie generatora liczb pseudolosowych.
    '2. Losuje 1 liczbę z przedziału od 0 do 10
    '   i przypisuje wylosowaną liczbę do zmiennej
    '   liczbowej całkowitej "PytA".
End Sub

Function ProcentOblicz(L As Integer, Lmaks As Integer) As String
'ProcentOblicz - Funkcja zwraca obliczony procent.
'
    If (Lmaks > 0) Then
        ProcentOblicz = " (" + CStr(Round(((L * 100) / Lmaks), 0)) + "%)"
    Else
        ProcentOblicz = " (0%)"
    End If
End Function

Sub Sprawdz()
'Sprawdz - Funkcja sprawdza poprawność odpowiedzi.
'
    Dim KomentarzP(4) As String
    Dim KomentarzN(4) As String
    Dim Ocena As Integer
    'Komentarze pozytywne.
        KomentarzP(0) = "Bardzo Dobrze"
        KomentarzP(1) = "Dobrze"
        KomentarzP(2) = "Doskonale"
        KomentarzP(3) = "Rewelacja"
        KomentarzP(4) = "Tak trzymać"
    'Komentarze negatywne.
        KomentarzN(0) = "Bardzo Źle"
        KomentarzN(1) = "Źle"
        KomentarzN(2) = "Błędna odpowiedź"
        KomentarzN(3) = "Niestety nie"
        KomentarzN(4) = "Zła odpowiedź"
    'Sprawdź odpowiedź.
        Randomize
        If (Trim(Cells(5, 8).Value) <> "") Then
            If ((PytA * PytB) = Cells(5, 8).Value) Then
                Cells(6, 8).Value = KomentarzP(Round((4 * Rnd), 0)) + "!"
                Call PytanieLosuj(False)
                OdpD = OdpD + 1
                Cells(7, 8).Value = ""
            Else
                Cells(6, 8).Value = KomentarzN(Round((4 * Rnd), 0)) + "!"
                Cells(7, 8).Value = "'Poprawna odpowiedź to: " + CStr(PytA) + " x " + CStr(PytB) + " = " + CStr(PytA * PytB)
                Call PytanieLosuj(True)
                OdpZ = OdpZ + 1
            End If
            Application.Wait (Now + TimeValue("00:00:02")) '1.
        End If
        Cells(5, 8).Value = ""
        Cells(6, 8).Value = ""
        Range("H5").Select
    'Pokaż punktacje.
        Cells(1, 4).Value = "'" + CStr(OdpD) + ProcentOblicz(OdpD, (OdpD + OdpZ))
        Cells(2, 4).Value = "'" + CStr(OdpZ) + ProcentOblicz(OdpZ, (OdpD + OdpZ))
    'Wystaw ocenę.
        Ocena = 0
        If ((OdpD + OdpZ) > 0) Then                 'Suma odpowiedzi pozytywnych i negatywnych musi być > 0.
            Ocena = (OdpD * 100 / (OdpD + OdpZ))    'Obliczenie, jaki procent stanowią odpowiedzi pozytywne.
        End If
        If (Ocena > 95) Then
            Cells(3, 4).Value = "Bardzo dobry (5.0)"
        ElseIf (Ocena > 80) Then
                Cells(3, 4).Value = "Dobry (4.0)"
            ElseIf (Ocena > 70) Then
                    Cells(3, 4).Value = "Dostateczny (3.0)"
                ElseIf (Ocena > 50) Then
                        Cells(3, 4).Value = "Dopuszczający (2.0)"
                    ElseIf (Ocena <= 50) Then
                            Cells(3, 4).Value = "Jedynka (1.0)"
                        Else
                            Cells(3, 4).Value = "?"
                        End If
    '
    'Legenda:
    '1. Zatrzymanie działania programu na 2 sekundy.
    '   Now - Zwraca wartość określającą bieżącą datę
    '         i godzinę zgodnie z datą i godziną systemową komputera.
    '   TimeValue(Parametr) - Konwertuje podany ciąg znaków na czas.
    '                         W miejsce parametru umieszcza się ciąg
    '                         znaków reprezentujący czas (np. "00:00:02").
End Sub

Sub StartTM()
'StartTM - Rozpoczyna program, przypisuje zmiennym wartości startowe, wywołuje poszczególne procedury lub funkcje.
'
    OdpD = 0
    OdpZ = 0
    'Pokaż punktacje.
        Cells(1, 2).Value = "Odp. dobra"
        Cells(1, 3).Value = ":"
        Cells(2, 2).Value = "Odp. zła"
        Cells(2, 3).Value = ":"
        Cells(3, 2).Value = "Ocena"
        Cells(3, 3).Value = ":"
        Cells(3, 4).Value = ""
        Cells(1, 4).Value = "'(0%)"
        Cells(2, 4).Value = "'(0%)"
        Cells(7, 8).Value = ""
    'Wywołanie procedur lub funkcji.
        Call PytanieLosuj(False)
End Sub