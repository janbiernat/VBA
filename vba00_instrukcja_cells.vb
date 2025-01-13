Sub instrukcja_cells()
'Instrukcja CELLS (c)by Jan T. Biernat
'
    Cells(3, 2).Value = "B3"                '1.
    Cells(5, 1).Value = Cells(3, 2).Value   '2.
    '
    'Legenda:
    '1. Umieszczenie w komórce B3 (tj. w wierszu nr 3 i kolumnie nr 2)
    '   tekstu znajdującego się pomiędzy cudzysłowami.
    '   Instrukcja CELLS służy do umieszczania danych w wybranej komórce.
    '
    '2. Przepisanie danych z komórki B3 do komórki A5.
    '   Komórka B3 znajduje się w wierszu nr 3 i kolumnie nr 2,
    '   natomiast komórka A5 znajduje się w wierszu nr 5 i kolumnie nr 1.
    '   Instrukcja CELLS służy również do pobierania danych z wybranej komórki.
End Sub
