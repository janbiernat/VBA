Sub wypisz_liczby_naturalne() 
'Wypisz liczby naturalne. 
'Copyright (c)by Jan T. Biernat 
' 
'Liczby naturalne - to liczby ca�kowite, dodatnie: 1,2,3,4,5,6, ... . 
'                   Czasami do liczb naturalnych zaliczamy r�wnie� liczb� zero. 
' 
    Dim I As Long                                                  '1 
    Cells(1, 1).Value = "Zakres:"                                  '2 
    Cells(1, 3).Value = ""                                         '3 
    If (Cells(1, 2).Value > 0) Then                                '4 
        For I = 1 To Cells(1, 2).Value                             '5 
            Cells(1 + I, 1).Value = I                              '6 
        Next I                                                     '5 
    Else                                                           '4 
        Cells(1, 3).Value = "B��D -?To nie jest liczba naturalna!" '7 
    End If                                                         '4 
    Cells(1, 2).Select 
    ' 
    'Legenda: 
    '1. Zadeklarowanie zmiennej "I" typu liczbowego ca�kowitego. 
    '   Typ Long jest wstanie przechowywa� liczby z zakresu od �2 147 483 648 
    '   do 2 147 486 647 (zajmuje 4B w pami�ci). 
    '   Wyraz Dim to skr�t od dimension - rozmiar, wymiar, wielko��. 
    '2. Wpisanie do kom�rki A1 (wiersz 1, kolumna 1 w arkuszu) tekstu znajduj�cego 
    '   si� pomi�dzy cudzys�owami. 
    '3. Wyczyszczenie kom�rki C1 (wiersz 1 i kolumna 3 w arkuszu). 
    '4. Sprawdzenie, czy w kom�rce B1 (wiersz 1 i kolumna 2 w arkuszu) podana warto�� jest wi�ksza od 0. 
    '   Je�eli tak, to warunek jest spe�niony. Gdy warunek jest spe�niony, to wykonaj instrukcje po s�owie THEN. 
    '   W przeciwnym przypadku wykonaj instrukcje po s�owie ELSE, czyli wy�wietl komunikat 
    '   "B��D -?To nie jest liczba naturalna!" 
    '5. Pocz�tek p�tli FOR. 
    '   For I = 1 To Cells(1, 2).Value - zapis ten okre�la ile razy p�tla b�dzie wykonana. 
    '                                    Wykonana b�dzie N razy. Zale�y to od warto�ci wpisanej 
    '                                    w kom�rce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '                                    R�wnie� instrukcje zawarte pomi�dzy FOR a NEXT te� b�d� 
    '                                    wykonane N razy. Czyli do kolejnych wierszy (licz�c 
    '                                    od 2 wiersza w 1 kolumnie w arkuszu) b�d� wpisywane 
    '                                    liczby z zakresu do 1 do N. W zale�no�ci, co zostanie wpisane 
    '                                    w kom�rce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '   Next I                         - koniec bloku p�tli FOR. Pomi�dzy FOR a NEXT umieszcza 
    '                                    si� instrukcj� lub blok instrukcji. 
    '6. Cells(2+I, 1).Value = I - wpisywanie/umieszczanie warto�ci liczbowej do kolejnych kom�rek 
    '                             w kolumnie 1 arkusza. Nr wiersza oraz warto�� kt�ra umieszczana 
    '                             jest w poszczeg�lnych kom�rkach, przechowywana jest w zmiennej liczbowej "I". 
    '                             Na pocz�tku zmienna liczbowa "I" przechowuje warto�� 1. 
    '                             Przy kolejnym powt�rzeniu p�tli warto�� przechowywana przez zmienn� "I" 
    '                             zamieni si� na liczb� 2. 
    '                             Przy kolejnym powt�rzeniu na 3 i tak do momentu gdy, przechowywana 
    '                             warto�� b�dzie mia�a liczb� N. 
    '                             Liczba N jest zale�na od wpisanej warto�ci w kom�rce B1 (wiersz 1 i kolumna 2 w arkuszu). 
    '                             Wtedy p�tla zako�czy swoje dzia�anie. 
    '7. Umieszczenie komunikatu "B��D -?To nie jest liczba naturalna!" w kom�rce C1 (wiersz 1 i kolumna 3 w arkuszu). 
    '   Instrukcja znajduje si� pomi�dzy konstrukcj� ELSE a END IF. 
End Sub 