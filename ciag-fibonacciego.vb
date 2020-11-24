Sub ciag_fibonacciego()
'Ciąg Fibonacciego by Leonardo Fibonacci
'Implementacja (c)by Jan T. Biernat
'
    Dim I As Byte                                                             '1
    Cells(1, 2).Value = 0                                                     '2
    Cells(2, 2).Value = 1                                                     '3
    For I = 0 To 37                                                           '4
        Cells(3 + I, 2).Value = Cells(1 + I, 2).Value + Cells(2 + I, 2).Value '5
    Next I                                                                    '4
    '
    'Legenda:
    '1. Zadeklarowanie zmiennej "I". Zmienna ta jest typu liczbowego całkowitego.
    '   Typ Byte zajmuje w pamięci 1B (bajt).
    '2. Wpisanie do komórki B1 (wiersz 1 i kolumna 2 w arkuszu) wartości 0.
    '3. Wpisanie do komórki B2 (wiersz 2 i kolumna 2 w arkuszu) wartości 1.
    '4. Pętla FOR zostanie wykonana 38 razy ponieważ zaczyna się
    '   od zera i musi osiągnąć wartość 37, co daje 38 powtórzeń.
    '5. Przepisanie do komórki B3 (wiersz 3 i kolumna 2 w arkuszu) wyniku
    '   sumowania dwóch liczb, które są pobrane z dwóch wcześniejszych
    '   komórek B1 i B2. Wszystko odbywa się w kolumnie 2 w arkuszu.
    '   Dzięki pętli FOR numeracja komórek (tj. numer wiersza - parametr 1) przy
    '   każdym powtórzeniu ulega zmianie (tj. zwiększeniu nr wiersza dzięki zmiennej "I").
    '   Wszystko odbywa się w kolumnie 2 w arkuszu.
End Sub