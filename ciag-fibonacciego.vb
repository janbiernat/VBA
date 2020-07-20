Sub ciag_fibonacciego()
'Ciąg Fibonacciego by Leonardo Fibonacci
'Oprogramowanie (c)by Jan T. Biernat
'
    A = 0 'Zadeklarowanie zmiennej liczbowej "A" i przypisanie do niej wartości 0.
    B = 1
    Cells(1, 2).Value = A 'Przepisanie zawartości zmiennej liczbowej "A" do komórki B1 (tj. w1, k2).
    Cells(2, 2).Value = B
    For I = 0 To 37 'Instrukcje w pętli FOR wykonają się 38 razy.
        C = 0     'Wyzerowanie zmiennej liczbowej "C".
        C = A + B 'Sumowanie liczb, które są przechowywane przez zmienne liczbowe "A" i "B"
                  'oraz przepisanie wyniku sumy do zmiennej liczbowej "C".
        Cells(3 + I, 2).Value = C 'Wypisanie wartości w komórkach (tj. w kolumnie 2 w poszczególnych wierszach)
        A = 0 'Wyzerowanie zmiennej liczbowej "A".
        A = B 'Przepisanie zawartości zmiennej liczbowej "B" do zmiennej liczbowej "A".
        B = 0
        B = C
    Next I
End Sub
