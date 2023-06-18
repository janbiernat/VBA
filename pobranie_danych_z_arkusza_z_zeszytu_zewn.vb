Sub KomorkaOdczyt()
'KomorkaOdczyt (c)by Jan T. Biernat
'
'Odczytanie danych z zewnętrznego zeszytu z wybranego arkusza i wybranej komórki.
'
    'Zapisanie w komórce B2 (w arkuszu np. "Arkusz1") danych, które są przypisane (pobrane) z zewnętrznego zeszytu (tj. np. "zeszyt1.xlsx") z arkusza (tj. np. "Arkusz1") i z wybranej komórki (tj. "A3").
        Sheets("Arkusz1").Range("B2") = Workbooks("zeszyt1.xlsx").Sheets("Arkusz1").Range("A3")
    'Zapisanie w komórce B3 (tj. wiersz nr 3, kol. nr 2 w arkuszu np. "Arkusz1") danych, które są przypisane (pobrane) z zewnętrznego zeszytu (tj. np. "zeszyt1.xlsx") z arkusza (tj. np. "Arkusz1") i z wybranej komórki A4 (tj. wiersz nr 4 i kol. nr 1).
        Sheets("Arkusz1").Cells(3, 2).Value = Workbooks("zeszyt1.xlsx").Sheets("Arkusz1").Cells(4, 1).Value
End Sub