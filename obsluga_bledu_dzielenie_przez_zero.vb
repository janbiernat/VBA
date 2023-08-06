Sub ObslugaBledu_DzieleniePrzezZero()
'ObslugaBledu_DzieleniePrzezZero - Obsługa błędu przez użytkownika.
'Copyright (c)by Jan T. Biernat
'
    'Wpisanie do komórek A1 i B1 wartości domyślnych.
        If ((Trim(Cells(1, 1).Value) = "") And (Trim(Cells(1, 2).Value) = "")) Then
            'Jeżeli komórki A1 i B1 są puste, to wykonaj poniższe instrukcje.
            Cells(1, 1).Value = 3
            Cells(1, 2).Value = 2
        End If
    'Obsługa błędu przez użytkownika.
        On Error Resume Next                                                            '1.
            Cells(1, 3).Value = Cells(1, 1).Value / Cells(1, 2).Value                   '2.
            If (Err = 11) Then                                                          '3.
                Cells(1, 3).Value = "BŁĄD -?Dzielenie przez zero jest niewykonalne!"    '4.
            End If                                                                      '3.
        On Error GoTo 0                                                                 '5.
'
'Legenda:
'1. On Error Resume Next
'   Konstrukcja ta nakazuje VBA na ignorowanie błędów w kodzie umieszczonym poniżej,
'   aż do momentu natrafienia na konstrukcję "On Error GoTo 0".
'
'2. Cells(1, 3).Value = Cells(1, 1).Value / Cells(1, 2).Value
'   Pobranie danych z komórek A1(w1,k1) i B1(w1,k2) oraz wykonanie dzielenia
'   na pobranych liczbach. Po wykonaniu działania, wynik zapisywany jest
'   do komórki C1(w1,k3).
'
'3. If (Err = 11) Then
'   Sprawdzenie za pomocą funkcji ERR wygenerowanego kodu błędu.
'   Jeżeli kod błędu będzie równy 11, to wyświetl komunikat o braku
'   możliwości wykonania dzielenia przez zero.
'4. Cells(1, 3).Value = "...
'   Wpisanie do komórki C1(w1,k3) informacji o błędzie
'   (tj. "BŁĄD -?Dzielenie przez zero jest niewykonalne!").
'5. On Error GoTo 0
'   Wyłączenie obsługi błędów przez użytkownika.
'
End Sub