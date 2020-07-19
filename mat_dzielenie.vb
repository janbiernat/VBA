Sub mat_dzielenie()
'Dzielenie (c)by Jan T. Biernat
'
    If (Cells(2, 1).Value = 0) Then
        Cells(3, 1).Value = "BŁĄD -?Dzielenie przez zero jest niewykonalne!"
    Else
        Cells(3, 1).Value = Cells(1, 1).Value / Cells(2, 1).Value
    End If
    '"Cells(2, 1).Value = 0" - Sprawdzenie, czy w komórce A2 (tj. w2, k1)
    '                          jest wartość zerowa.
    'Jeżeli w komórce A2 (tj. w2, k1) jest wartość zerowa, to wykonaj
    'instrukcje po słowie THEN. Przeciwnym razie wykonaj instrukcje
    'po słowie ELSE (tj. wykonaj działanie dzielenia).
    '
    'Cells(nr_wiersza, nr_kolumny).Value - Umożliwia pobranie danych z komórki
    '                                      lub wprowadzenie danych do komórki
    '                                      o nr wiersza "nr_wiersza"
    '                                      i nr kolumny "nr_kolumny".
End Sub