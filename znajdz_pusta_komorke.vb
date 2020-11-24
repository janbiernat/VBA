Function ZnajdzPustaKomorke(ArkuszNazwa As String, KolumnaNr As Long)
'Znajdź pustą komórkę.
'Copyright (c)by Jan T. Biernat
'
'Wywołanie funkcji:
'zmienna_liczbowa = ZnajdzPustaKomorke(zmienna_tekstowa, zmienna_liczbowa_nr_kolumny)
'
    Dim Licznik As Long                                    'Zadeklarowanie zmiennej lokalnej liczbowej całkowitej "Licznik". Typ Long może przechować liczby całkowite z zakresu od –2 147 483 648 do 2 147 486 647 i zajmuje 4B pamięci.
    Licznik = 0                                            'Przypisanie zmiennej liczbowej "Licznik" wartość 0.
    With Sheets(ArkuszNazwa)                               'Początek konstrukcji wiążącej With ... End With. Wewnątrz tej konstrukcji wszystkie instrukcje można używać bez przedrostka umieszczonego po słowie With.
        Do                                                 'Początek pętli DO ... LOOP UNTIL (warunek). Wewnątrz tej pętli umieszczone instrukcje wykonywane są tak długo, aż warunek zostanie spełniony.
            Licznik = Licznik + 1                          'Zwiększenie wartości liczbowej (przechowywanej przez zmienną liczbową "Licznik") o wartość 1. To jest inkrementacja.
        Loop Until (.Cells(Licznik, KolumnaNr).Value = "") 'Koniec pętli DO ... LOOP UNTIL (warunek). Zakończenie działania pętli następuje po spełnieniu warunku. Warunek umieszczony jest na końcu po słowie UNTIL.
    End With                                               'Zakończenie konstrukcji wiążącej With ... End With.
    ZnajdzPustaKomorke = Licznik                           'Zwrócenie wyniku działania funkcji ZnajdzPustaKomorke do miejsca z którego funkcja została wywołana.
End Function

Sub znajdz_pusta_komorke()
'Wyszukaj pustą komórkę
'Copyright (c)by Jan T. Biernat
'
    Dim W_K1 As Long
    Dim W_K2 As Long
    Dim W_K3 As Long
    W_K1 = 0
    W_K1 = ZnajdzPustaKomorke("Pętla DO ... LOOP UNTIL", 1) 'WWywołanie funkcji ZnajdzPustaKomorke oraz przepisanie wyniku działania funkcji do zmiennej liczbowej "W_K1".
    Cells(W_K1, 1).Value = W_K1
    W_K2 = 0
    W_K2 = ZnajdzPustaKomorke("Pętla DO ... LOOP UNTIL", 2)
    Cells(W_K2, 2).Value = W_K2
    W_K3 = 0
    W_K3 = ZnajdzPustaKomorke("Pętla DO ... LOOP UNTIL", 3)
    Cells(W_K3, 3).Value = W_K3
End Sub