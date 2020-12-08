Sub wypisz_liczby_naturalne()
'Wypisz liczby naturalne
'Copyright (c)by Jan T. Biernat
'
'Liczby naturalne - to liczby całkowite, dodatnie: 1,2,3,4,5,6, ... .
'                   Czasami do liczb naturalnych zaliczamy również liczbę zero.
'
    Dim I As Long                                                  '1
    Cells(1, 1).Value = "Zakres:"                                  '2
    Cells(1, 3).Value = ""                                         '3
    If (Cells(1, 2).Value > 0) Then                                '4
        For I = 1 To Cells(1, 2).Value                             '5
            Cells(1 + I, 1).Value = I                              '6
        Next I                                                     '5
    Else                                                           '4
        Cells(1, 3).Value = "BŁĄD -?To nie jest liczba naturalna!" '7
    End If                                                         '4
    Cells(1, 2).Select
    '
    'Legenda:
    '1. Zadeklarowanie zmiennej "I" typu liczbowego całkowitego.
    '   Typ Long jest wstanie przechowywać liczby z zakresu od –2 147 483 648
    '   do 2 147 486 647 (zajmuje 4B w pamięci).
    '   Wyraz Dim to skrót od dimension - rozmiar, wymiar, wielkość.
    '2. Wpisanie do komórki A1 (wiersz 1, kolumna 1 w arkuszu) tekstu znajdującego
    '   się pomiędzy cudzysłowami.
    '3. Wyczyszczenie komórki C1 (wiersz 1 i kolumna 3 w arkuszu).
    '4. Sprawdzenie, czy w komórce B1 (wiersz 1 i kolumna 2 w arkuszu) podana wartość jest większa od 0.
    '   Jeżeli tak, to warunek jest spełniony. Gdy warunek jest spełniony, to wykonaj instrukcje po słowie THEN.
    '   W przeciwnym przypadku wykonaj instrukcje po słowie ELSE, czyli wyświetl komunikat
    '   "BŁĄD -?To nie jest liczba naturalna!"
    '5. Początek pętli FOR.
    '   For I = 1 To Cells(1, 2).Value - zapis ten określa ile razy pętla będzie wykonana.
    '                                    Wykonana będzie N razy. Zależy to od wartości wpisanej
    '                                    w komórce B1 (wiersz 1 i kolumna 2 w arkuszu).
    '                                    Również instrukcje zawarte pomiędzy FOR a NEXT też będą
    '                                    wykonane N razy. Czyli do kolejnych wierszy (licząc
    '                                    od 2 wiersza w 1 kolumnie w arkuszu) będą wpisywane
    '                                    liczby z zakresu do 1 do N. W zależności, co zostanie wpisane
    '                                    w komórce B1 (wiersz 1 i kolumna 2 w arkuszu).
    '   Next I                         - koniec bloku pętli FOR. Pomiędzy FOR a NEXT umieszcza
    '                                    się instrukcję lub blok instrukcji.
    '6. Cells(2+I, 1).Value = I - wpisywanie/umieszczanie wartości liczbowej do kolejnych komórek
    '                             w kolumnie 1 arkusza. Nr wiersza oraz wartość która umieszczana
    '                             jest w poszczególnych komórkach, przechowywana jest w zmiennej liczbowej "I".
    '                             Na początku zmienna liczbowa "I" przechowuje wartość 1.
    '                             Przy kolejnym powtórzeniu pętli wartość przechowywana przez zmienną "I"
    '                             zamieni się na liczbę 2.
    '                             Przy kolejnym powtórzeniu na 3 i tak do momentu gdy, przechowywana
    '                             wartość będzie miała liczbę N.
    '                             Liczba N jest zależna od wpisanej wartości w komórce B1 (wiersz 1 i kolumna 2 w arkuszu).
    '                             Wtedy pętla zakończy swoje działanie.
    '7. Umieszczenie komunikatu "BŁĄD -?To nie jest liczba naturalna!" w komórce C1 (wiersz 1 i kolumna 3 w arkuszu).
    '   Instrukcja znajduje się pomiędzy konstrukcją ELSE a END IF.
End Sub
