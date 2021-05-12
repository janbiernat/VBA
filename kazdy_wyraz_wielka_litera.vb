Sub kazdy_wyraz_wielka_litera() 
'Każdy wyraz wielką literą. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim WierszNr As Integer                                                                                                                                    '1. 
    Dim WierszT As String                                                                                                                                      '2. 
    Dim I As Integer                                                                                                                                           '3. 
    Dim WierszWynik As String                                                                                                                                  '4. 
    WierszNr = 0                                                                                                                                               '5. 
    With Sheets("Każdy wyraz wielką literą")                                                                                                                   '6. 
        Do                                                                                                                                                     '7. 
            WierszNr = WierszNr + 1                                                                                                                            '8. 
            WierszT = ""                                                                                                                                       '9. 
            WierszT = LCase(.Cells(WierszNr, 1).Value)                                                                                                         '10. 
            WierszWynik = ""                                                                                                                                   '11. 
            For I = 2 To Len(WierszT)                                                                                                                          '12. 
                If (((Mid(WierszT, I - 1, 1) = " ") And (Mid(WierszT, I, 1) <> " ")) Or ((Mid(WierszT, I - 1, 1) = "-") And (Mid(WierszT, I, 1) <> " "))) Then '13. 
                    WierszWynik = WierszWynik + UCase(Mid(WierszT, I, 1))                                                                                      '14. 
                Else                                                                                                                                           '13. 
                    WierszWynik = WierszWynik + Mid(WierszT, I, 1)                                                                                             '15. 
                End If                                                                                                                                         '13. 
            Next I                                                                                                                                             '12. 
            .Cells(WierszNr, 1).Value = UCase(Mid(WierszT, 1, 1)) + WierszWynik                                                                                '16. 
        Loop Until (.Cells(WierszNr, 1).Value = "")                                                                                                            '7. 
    End With                                                                                                                                                   '6. 
End Sub
'
'Legenda: 
'  1) Zadeklarowanie zmiennej "WierszNr", która jest typu liczbowego całkowitego. 
'     Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'     z zakresu od -2 147 483 648 do 2 147 483 647. 
'  2) Zadeklarowanie zmiennej "WierszT", która jest typu tekstowego. 
'     Zmienna tego typu umożliwia przechowywanie znaków alfanumerycznych. 
'  3) Zadeklarowanie zmiennej "I", która jest typu liczbowego całkowitego. 
'     Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'     z zakresu od -2 147 483 648 do 2 147 483 647.
'  4) Zadeklarowanie zmiennej "WierszWynik", która jest typu tekstowego. 
'     Zmienna tego typu umożliwia przechowywanie znaków alfanumerycznych. 
'  5) Przypisanie do zmiennej liczbowej "WierszNr" wartości zerowej. 
'  6) Instrukcja wiążąca WITH ... END WITH umożliwia używanie innych instrukcji 
'     bez potrzeby pisania pełniej składni. Pod warunkiem, że te instrukcje znajdują 
'     się pomiędzy WITH a END WITH. 
'     Na przykład: 
'                  Dzięki instrukcji wiążącej wystarczy napisać ".Cells(WierszNr, 1).Value" 
'                  zamiast "Sheets("Każdy wyraz wielką literą").Cells(WierszNr, 1).Value". 
'  7) Pętla DO ... LOOP UNTIL(warunek). 
'     Pętla ta wykonuje instrukcje w niej zawarte określoną ilość razy. 
'     Ilość powtórzeń instrukcji w pętli zależy od warunku, który umieszczony jest na końcu 
'     pętli (za instrukcją UNTIL). Umieszczenie warunku na końcu pętli powoduje, że instrukcje 
'     w niej zawarte będą wykonane zawsze raz. 
'     Pętla DO ... LOOP UNTIL(warunek) zakończy swoje działanie w momencie natrafienia na pustą komórkę. 
'  8) WierszNr = WierszNr + 1. 
'     Jest to zwiększenie (tzw. inkrementacja) zawartości zmiennej liczbowej całkowitej 
'     "WierszNr" o wartość 1. 
'  9) Wyczyszczenie zmiennej tekstowej "WierszT". 
' 10) Przypisanie do zmiennej tekstowej "WierszT" zawartości komórki, która znajduje się w kolumnie nr 1 (w arkuszu to kolumna A) 
'     w wierszu o nr przechowywanym w zmiennej liczbowej całkowitej "WierszNr". 
'     LCase(P) - Zamienia cały ciąg znaków na małe litery. 
'                W miejsce "P" należy umieścić zmienna tekstową (np. ".Cells(WierszNr, 1).Value"). 
' 11) Wyczyszczenie zmiennej tekstowej "WierszWynik". 
' 12) Pętla FOR ... NEXT. 
'     Pętla wykonuje instrukcje w niej zawarte określoną ilość razy. Ilość powtórzeń jest określona 
'     za pomocą konstrukcji Len(WierszT). 
'     Len(P) - Instrukcja zwraca liczbę całkowitą określającą z ilu znaków składa się podany tekst. 
'              W miejsce parametru P umieszcza się zmienną tekstową przechowującą jakiś ciąg znaków. 
' 13) IF (warunek) THEN. 
'     Instrukcja warunkowa, której zadaniem jest sprawdzenie poprawności podanego warunku lub warunków. 
'     Warunek 1 - jest to warunek, który składa się z dwóch warunków. 
'                 W tym warunku jest szukany znak spacji [tj. "(Mid(WierszT, I - 1, 1) = " ")"] 
'                 i dowolny znak za wyjątkiem spacji [tj "(Mid(WierszT, I, 1) <> " ")"]. 
'     Warunek 2 - jest to warunek, który składa się z dwóch warunków. 
'                 W tym warunku jest szukany znak "-" [tj. "(Mid(WierszT, I - 1, 1) = "-")"] 
'                 i dowolny znak za wyjątkiem spacji [tj "(Mid(WierszT, I, 1) <> " ")"]. 
'     Wyszukanie obydwóch znaków odbywa się w podanym tekście. 
'     Jeżeli jeden z dwóch konstrukcji warunków połączonych będzie spełniony, to wykonaj instrukcje po słowie THEN. 
'     MID(Tekst, Start, Ile) - Wyciąga fragment tekstu z podanego ciągu znaków. 
'                              W miejsce parametru "Tekst" umieszcza się zmienną tekstową przechowującą dowolny tekst. 
'                              W miejsce parametru "Start" umieszcza się wartość liczbową całkowitą określającą początek pobierania fragmentu tekstu. 
'                              W miejsce parametru "Ile" umieszcza się wartość liczbową całkowitą określającą ile znaków ma zostać pobranych. 
' 14) WierszWynik = WierszWynik + UCase(Mid(WierszT, I, 1)). 
'     Dodanie do zmiennej tekstowej "WierszWynik" fragmentu tekstu, który został pobrany za pomocą instrukcji MID. 
'     Instrukcja ta wyciąga znak, który znajduje się na pozycji o nr przechowywanym w zmiennej liczbowej "I". 
'     UCase(P) - Przekształca wszystkie znaki podanego ciągu tekstowego na wielkie litery (np. Atari -> ATARI).  
' 15) WierszWynik = WierszWynik + Mid(WierszT, I, 1). 
'     Dodanie do zmiennej tekstowej "WierszWynik" fragmentu tekstu, który został pobrany za pomocą instrukcji MID. 
'     Instrukcja ta wyciąga znak, który znajduje się na pozycji o nr przechowywanym w zmiennej liczbowej "I". 
' 16) .Cells(WierszNr, 1).Value = UCase(Mid(WierszT, 1, 1)) + WierszWynik 
'     Przypisanie do komórki w kolumnie nr 1 (w arkuszu to kolumna A), w wierszu o nr przechowywanym w zmiennej liczbowej "WierszNr" 
'     pewnej wartości. Wartość ta, to 1 litera pobranego ciągu znaków z danej komórki + zawartość zmiennej "WierszWynik". 