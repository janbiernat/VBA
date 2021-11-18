Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
'Zdarzenia "Worksheet_SelectionChange" używa się w celu 
'wykonania zaprogramowanego(określonego) zadania, 
'po dokonanej zmianie w wybranej komórce. 
'W tym przykładzie zdarzenie "Worksheet_SelectionChange" zostanie wykorzystanie 
'do obliczenia wyniku z dzielenia. 
' 
    If ((Cells(1, 1).Value <> "") And (Cells(2, 1).Value <> "")) Then            '1 
        If (Cells(2, 1).Value = 0) Then                                          '2 
            Cells(3, 1).Value = "BŁĄD -?Dzielenie przez zero jest niewykonalne!" '3 
        Else                                                                     '2 
            Cells(3, 1).Value = Cells(1, 1).Value / Cells(2, 1).Value            '4 
        End If                                                                   '2 
    Else                                                                         '1 
        Cells(3, 1).Value = "INFO: Komórka A1 i A2 musi zawierać liczbę!"        '5 
    End If                                                                       '1 
    ' 
    'Legenda: 
    '1) Sprawdzenie, czy w komórkach A1 i A2 są wpisane wartości liczbowe. 
    '   Do sprawdzenia tego warunku została użyta instrukcja IF, która zawiera 
    '   dwa warunki połączone operatorem AND. 
    '   Warunek 1: sprawdza, czy w komórce A1 (tj. w arkuszu komórka w kolumnie 1 i w wierszu 1) 
    '              jest wpisana wartość liczbowa. 
    '   Warunek 2: sprawdza, czy w komórce A2 (tj. w arkuszu komórka w kolumnie 1 i w wierszu 2) 
    '              jest wpisana wartość liczbowa. 
    '2) Sprawdzenie, czy w komórce A2 (tj. w arkuszu komórka w kolumnie 1 i w wierszu 2) 
    '   ma wpisaną wartość zerową. Jeżeli tak, to wyświetl komunikat 
    '   "BŁĄD -?Dzielenie przez zero jest niewykonalne!". W innym przypadku wykonaj instrukcje 
    '   po słowie ELSE (tj. wykonaj dzielenie liczby przez liczbę). 
    '3) Wyświetlenie komunikatu o niemożności wykonania dzielenia liczby przez liczbę. 
    '4) Wykonanie dzielenia liczby przez liczbę. Pierwsza liczba jest w komórce A1 
    '   (tj. w arkuszu komórka w kolumnie 1 i w wierszu 1), 
    '   natomiast druga liczba jest w komórce A2 (tj. w arkuszu komórka w kolumnie 1 i w wierszu 2). 
    '   Natomiast w komórce A3 (tj. w arkuszu w kolumnie 1 i w wierszu 3) pojawia się wynik 
    '   z dzielenia dwóch liczb. 
    '5) Wyświetlenie komunikatu w komórce A3 (tj. w arkuszu w kolumnie 1 i wierszu 3) 
    '   o braku wpisanych liczb w komórkach A1 i A2. 
End Sub 