Sub wypisz_liczby_naturalne() 
'Wypisz liczby naturalne. 
'Copyright (c)by Jan T. Biernat 
' 
'Liczby naturalne - to liczby ca³kowite, dodatnie: 1,2,3,4,5,6, ... . 
'                   Czasami do liczb naturalnych zaliczamy równie¿ liczbê zero. 
' 
    Dim I As Long                                                  '1 
    Cells(1, 1).Value = "Zakres:"                                  '2 
    Cells(1, 3).Value = ""                                         '3 
    If (Cells(1, 2).Value > 0) Then                                '4 
        For I = 1 To Cells(1, 2).Value                             '5 
            Cells(1 + I, 1).Value = I                              '6 
        Next I                                                     '5 
    Else                                                           '4 
        Cells(1, 3).Value = "B£¥D -?To nie jest liczba naturalna!" '7 
    End If                                                         '4 
    Cells(1, 2).Select 
    ' 
    'Legenda: 
    '1. Zadeklarowanie zmiennej "I" typu liczbowego ca³kowitego. 
    '   Typ Long jest wstanie przechowywaæ liczby z zakresu od –2 147 483 648 
    '   do 2 147 486 647 (zajmuje 4B w pamiêci). 
    '   Wyraz Dim to skrót od dimension - rozmiar, wymiar, wielkoœæ. 
    '2. Wpisanie do komórki A1 (wiersz 1, kolumna 1 w arkuszu) tekstu znajduj¹cego 
    '   siê pomiêdzy cudzys³owami. 
    '3. Wyczyszczenie komórki C1 (wiersz 1 i kolumna 3 w arkuszu). 
    '4. Sprawdzenie, czy w komórce B1 (wiersz 1 i kolumna 2 w arkuszu) podana wartoœæ jest wiêksza od 0. 
    '   Je¿eli tak, to warunek jest spe³niony. Gdy warunek jest spe³niony, to wykonaj instrukcje po s³owie THEN. 
    '   W przeciwnym przypadku wykonaj instrukcje po s³owie ELSE, czyli wyœwietl komunikat 
    '   "B£¥D -?To nie jest liczba naturalna!" 
    '5. Pocz¹tek pêtli FOR. 
    '   For I = 1 To Cells(1, 2).Value - zapis ten okreœla ile razy pêtla bêdzie wykonana. 
    '                                    Wykonana bêdzie N razy. Zale¿y to od wartoœci wpisanej 
    '                                    w komórce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '                                    Równie¿ instrukcje zawarte pomiêdzy FOR a NEXT te¿ bêd¹ 
    '                                    wykonane N razy. Czyli do kolejnych wierszy (licz¹c 
    '                                    od 2 wiersza w 1 kolumnie w arkuszu) bêd¹ wpisywane 
    '                                    liczby z zakresu do 1 do N. W zale¿noœci, co zostanie wpisane 
    '                                    w komórce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '   Next I                         - koniec bloku pêtli FOR. Pomiêdzy FOR a NEXT umieszcza 
    '                                    siê instrukcjê lub blok instrukcji. 
    '6. Cells(2+I, 1).Value = I - wpisywanie/umieszczanie wartoœci liczbowej do kolejnych komórek 
    '                             w kolumnie 1 arkusza. Nr wiersza oraz wartoœæ która umieszczana 
    '                             jest w poszczególnych komórkach, przechowywana jest w zmiennej liczbowej "I". 
    '                             Na pocz¹tku zmienna liczbowa "I" przechowuje wartoœæ 1. 
    '                             Przy kolejnym powtórzeniu pêtli wartoœæ przechowywana przez zmienn¹ "I" 
    '                             zamieni siê na liczbê 2. 
    '                             Przy kolejnym powtórzeniu na 3 i tak do momentu gdy, przechowywana 
    '                             wartoœæ bêdzie mia³a liczbê N. 
    '                             Liczba N jest zale¿na od wpisanej wartoœci w komórce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '                             Wtedy pêtla zakoñczy swoje dzia³anie. 
    '7. Umieszczenie komunikatu "B£¥D -?To nie jest liczba naturalna!" w komórce C1 (wiersz 1 i kolumna 3 w arkuszu). 
    '   Instrukcja znajduje siê pomiêdzy konstrukcj¹ ELSE a END IF. 
End Sub 