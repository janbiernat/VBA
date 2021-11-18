Sub dzielenie_liczby_przez_zero() 
'Dzielenie liczby przez zero. 
'Copyright (c)by Jan T. Biernat 
' 
    If (Cells(2, 1).Value = 0) Then                                          '1. 
        Cells(3, 1).Value = "BŁĄD -?Dzielenie przez zero jest niewykonalne!" '2. 
    Else                                                                     '1. 
        Cells(3, 1).Value = Cells(1, 1).Value / Cells(2, 1).Value            '3. 
    End If                                                                   '1. 
    ' 
    'Legenda: 
    '1. Sprawdzenie, czy w komórce A2 (wiersz 2 i kolumna 1 w arkuszu) znajduje się 
    '   liczba 0. Jeżeli tak, to wyświetl komunikat w komórce A3 (wiersz 3 i kolumna 1 w arkuszu) 
    '   o treści "BŁĄD -?Dzielenie przez zero jest niewykonalne!". 
    '   W innym przypadku wykonaj działanie dzielenia liczby przez inną liczbę 
    '   tzn. podziel zawartość komórki A1 (wiersz 1 i kolumna 1 w arkuszu) przez zawartość 
    '   komórki A2 (wiersz 2 i kolumna 1 w arkuszu), co spowoduje wyświetlenie wyniku 
    '   w komórce A3 (wiersz 3 i kolumna 1 w arkuszu). 
    '   Jeżeli warunek jest prawdziwy to wykonaj instrukcje po słowie THEN, 
    '   czyli wyświetl komunikat o błędzie. W innym przypadku wykonaj instrukcje 
    '   po słowie ELSE, czyli wykonaj dzielenie liczby przez inną liczbę. 
    '2. Wyświetlenie w komórce A3 (wiersz 3 i kolumna 1 w arkuszu) komunikatu o treści 
    '   "BŁĄD -?Dzielenie przez zero jest niewykonalne!". 
    '3. Wykonanie działania dzielenie liczby przez inną liczbę, 
    '   czyli liczby w komórce A1 przez liczbę w komórce A2. 
    '   Wynik będzie zapisany w komórce A3. 
End Sub 