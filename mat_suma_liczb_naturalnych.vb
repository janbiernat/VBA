Sub suma_liczb_naturalnych()
'Pętla FOR.
'Suma liczb naturalnych.
'Copyright (c)by Jan T. Biernat
'=
'Liczby naturalne - to liczby całkowite, dodatnie: 1, 2, 3, 4, 5, 6, ... .
'                   Czasami do liczb naturalnych zaliczamy również liczbę zero.
'
    Dim Suma As Integer             '1
    Dim I As Integer                '2
    Suma = 0                        '3
    Cells(1, 1).Value = "Liczba:"   '4
    If (Cells(1, 2).Value < 1) Then '5
        Cells(1, 2).Value = 1       '6
    End If                          '5
    For I = 1 To Cells(1, 2).Value  '7
        Suma = Suma + I             '8
        Cells(2, I).Value = I       '9
    Next I                          '7
    Cells(2, I).Value = "="         '10
    Cells(2, I + 1).Value = Suma    '11
    Cells(1, 2).Select              '12
    '
    'Legenda:
    ' 1. Zadeklarowanie zmiennej liczbowej całkowitej "Suma".
    '    Typ Integer jest wstanie przechować liczby z zakresu 
    '    od -32768 do 32767, co zajmuje 2 Bajty pamięci RAM.
    ' 2. Zadeklarowanie zmiennej liczbowej całkowitej "I".
    '    Typ Integer jest wstanie przechować liczby z zakresu 
    '    od -32768 do 32767, co zajmuje 2 Bajty pamięci RAM.
    ' 3. Przypisanie wartości początkowej 0 do zmiennej liczbowej "Suma".
    ' 4. Wpisanie do komórki A1 (wiersz 1, kolumna 1 w arkuszu) tekstu,
    '    który znajduje się pomiędzy cudzysłowami (tj. "Liczba:").
    ' 5. Sprawdzenie, czy w komórce B1 (wiersz 1 i kolumna 2 w arkuszu)
    '    podana wartość jest mniejsza od 1. Jeżeli tak, to warunek jest
    '    spełniony. Gdy warunek jest spełniony, to wykonaj instrukcje po
    '    słowie THEN (tj. wprowadź wartość domyślną 1).
    ' 6. Wpisanie do komórki B1 (wiersz 1, kolumna 2 w arkuszu) domyślnej wartości 1.
    ' 7. Ilość powtórzeń pętli FOR zależy od wartości, która znajduje się
    '    w komórce B1 (wiersz 1, kolumna 2 w arkuszu).
    '    Pętla FOR umożliwia powtarzanie instrukcji lub bloku instrukcji określoną
    '    wcześniej ilość razy. 
    '    Konstrukcja pętli FOR:
    '       FOR Zmienna = wartosc_poczatkowa TO wartosc_koncowa
    '           Instrukcja 1
    '           Instrukcja 2
    '           Instrukcja ...
    '           Instrukcja N
    '       NEXT Zmienna
    ' 8. Suma = Suma + I - Zwiększenie wartości przechowywanej przez zmienną liczbową "Suma"
    '                     o wartość, która jest przechowywana w zmiennej liczbowej "I".
    '    Zapis "Suma = Suma + 1" - jest to zwiększenie wartości przechowywanej w zmiennej
    '                              liczbowej "Suma" o wartość np. 1. Jest to tzn. inkrementacja.
    ' 9. Wpisanie do kolejnych komórek (licząc od kolumny 1, wiersz 2 w arkuszu) wartości
    '    przechowywanej w zmiennej liczbowej "I". Do zmiennej liczbowej "I" wartości są wpisywane
    '    przez pętlę FOR.
    '10. Wpisanie/umieszczenie tekstu znajdującego się pomiędzy cudzysłowami (tj. znaku "=")
    '    w wierszu 2 arkusza w kolumnie której nr przechowywany jest w zmiennej liczbowej "I".
    '11. Wpisanie/umieszczenie zawartości zmiennej liczbowej "Suma" w wierszu 2 arkusza
    '    w kolumnie której nr przechowywany jest w zmiennej liczbowej "I+1" powiększonej o wartość 1.
    '12. Zaznaczenie komórki B1 (wiersz 1, kolumna 2 w arkuszu).
End Sub