Private Sub Workbook_Open()
'Tabliczka mnożenia - Losowanie 1 kolumny z 10 kolumn.
'Copyright (c)by Jan T. Biernat
'
    'Deklaracja zmiennych.
    Dim A As Integer
    Dim I As Integer
    'Wylosowanie nr kolumny.
    Randomize               'Zainicjowanie generatora liczb losowych.
    A = 0                   'Przypisanie wartości zerowej do zmiennej liczbowej "A".
    A = Int((10 * Rnd) + 1) 'Generuje losową liczbę z zakresu od 1 do 10 i przypisuje zmiennej liczbowej "A".
    'Wyświetl wylosowaną kolumnę na ekranie.
    With Sheets("Tabliczka mnożenia - Losowanie")
        .Cells(1, 1).Value = "Kolumna nr: " + CStr(A)
        For I = 1 To 10
            Cells(2 + I, 1).Value = I
            Cells(2 + I, 2).Value = " * "
            Cells(2 + I, 3).Value = A
            Cells(2 + I, 4).Value = " = "
            Cells(2 + I, 5).Value = (I * A)
        Next I
    End With
End Sub
