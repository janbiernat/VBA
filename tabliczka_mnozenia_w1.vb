Sub tabliczka_mnozenia_w1()
'Tabliczka mnożenia w1 (c)by Jan T. Biernat
'
    Dim A As Integer
    Dim B As Integer
    For A = 1 To 10
        Cells(1, 1 + A).Value = A 'Wypisanie liczb od 1 do 10 w wierszu 1 zaczynając od kolumny 2.
        Cells(1 + A, 1).Value = A 'Wypisanie liczb od 1 do 10 w kolumnie 1 zaczynając od wiersza 2.
        For B = 1 To 10
            Cells(1 + B, 1 + A).Value = A * B 'Wyświetlenie wyniku z pomnożenia dwóch liczb znajdujących się
                                              'w wierszu nagłówkowym i w kolumnie nagłówkowej.
        Next B
    Next A
End Sub
