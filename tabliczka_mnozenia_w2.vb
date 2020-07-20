Sub tabliczka_mnozenia_w2()
'Tabliczka mnożenia w2 (c)by Jan T. Biernat
'
    Dim A As Integer
    Dim B As Integer
    Cells(1, 1).Value = "Zakres:" 'Wprowadź do komórki A1 (tj. w1, k1) napis "Zakres:".
    If (Cells(1, 2).Value > 1) Then
        For A = 1 To Cells(1, 2).Value
            Cells(6, 1 + A).Value = A 'Wyświetl nagłówek poziomy od 1 do N.
            Cells(6 + A, 1).Value = A 'Wyświetl nagłówek pionowy od 1 do N.
            For B = 1 To Cells(1, 2).Value
                Cells(6 + A, 1 + B).Value = (A * B) 'Przemnożenie na krzyż liczb w nagłówku poziomym (wiersz 6) z nagłówkiem pionowym (kolumna 1).
            Next B
        Next A
        Cells(2, 2).Value = "Wszystko jest w porządku!"
    Else
        Cells(2, 2).Value = "Błąd -?Zakres jest zbyt mały. Proszę wprowadzić minimum wartość 2."
    End If
    ActiveSheet.Cells(1, 2).Select 'lub Sheets("Tabliczka mnożenia w2").Cells(1, 2).Select
End Sub
