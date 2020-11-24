Sub wypisz_liczby_naturalne()
'Wypisz liczby naturalne
'Copyright (c)by Jan T. Biernat
'
'Liczby naturalne - to liczby całkowite, dodatnie: 1,2,3,4,5,6, ... .
'Czasami do liczb naturalnych zaliczamy również liczbę zero.
'
    Dim I As Long                                                  '1
    Cells(2, 3).Value = ""                                         '2
    If (Cells(2, 2).Value > 0) Then                                '3
        For I = 1 To Cells(2, 2).Value                             '4
            Cells(2 + I, 1).Value = I                              '5
        Next I                                                     '4
    Else                                                           '3
        Cells(2, 3).Value = "BŁĄD -?To nie jest liczba naturalna!" '6
    End If                                                         '3
    '
    'Legenda:
    '1. Zadeklarowanie zmiennej "I" typu liczbowego całkowitego.
    '   Typ Long jest wstanie przechowywać liczby z zakresu od –2 147 483 648
    '   do 2 147 486 647 (zajmuje 4B w pamięci).
    '   Wyraz Dim to skrót od dimension - rozmiar, wymiar, wielkość.
    '2. Wyczyszczenie komórki C2 (wiersz 2 i kolumna 3 w arkuszu).
    '3. Sprawdzenie, czy w komórce B2 (wiersz 2 i kolumna 2 w arkuszu) podana wartość jest większa od -1.
    '   Jeżeli tak, to warunek jest spełniony. Gdy warunek jest spełniony, to wykonaj instrukcje po słowie THEN.
    '   W przeciwnym przypadku wykonaj instrukcje po słowie ELSE, czyli wyświetl komunikat "BŁĄD -?To nie jest liczba naturalna!"
    '4. Początek pętli FOR.
    '   For I = 1 To Cells(2, 2).Value - zapis ten określa ile razy pętla będzie wykonana.
    '                                    Wykonana będzie N razy. Zależy to od wpisanej wartości
    '                                    w komórce B2 (wiersz 2 i kolumna 2 w arkuszu).
    '                                    Również instrukcje zawarte pomiędzy FOR a NEXT też będą wykonane N razy.
    '                                    Czyli do kolejnych wierszy (licząc od 3 wiersza w 1 kolumnie w arkuszu) będą
    '                                    wpisywane liczby z zakresu do 1 do N. W zależności, co zostanie wpisane
    '                                    w komórce B2 (wiersz 2 i kolumna 2 w arkuszu).
    '   Next I                         - koniec bloku pętli FOR. Pomiędzy FOR a NEXT umieszcza się instrukcję lub blok instrukcji.
    '5. Cells(2+I, 1).Value = I - wpisywanie/umieszczanie wartości liczbowej do kolejnych komórek w kolumnie 1 arkusza.
    '                             Nr wiersza oraz wartość która umieszczana jest w poszczególnych komórkach,
    '                             przechowywana jest w zmiennej liczbowej "I".
    '                             Na początku zmienna liczbowa "I" przechowuje wartość 1.
    '                             Przy kolejnym powtórzeniu pętli wartość przechowywana przez zmienną "I" zamieni się na liczbę 2.
    '                             Przy kolejnym powtórzeniu na 3 i tak do momentu gdy, przechowywana wartość będzie miała liczbę N.
    '                             Liczba N jest zależna od wpisanej wartości w komórce B2 (wiersz 2 i kolumna 2 w arkuszu).
    '                             Wtedy pętla zakończy swoje działanie.
    '6. Umieszczenie komunikatu "BŁĄD -?To nie jest liczba naturalna!" w komórce C2 (wiersz 2 i kolumna 3 w arkuszu).
    '   Instrukcja znajduje się pomiędzy konstrukcją ELSE a END IF.
End Sub