Sub inicjaly() 
'Inicjały (c)by Jan T. Biernat 
' 
    Dim WierszNr As Integer                                                '1. 
    Dim I As Integer                                                       '2. 
    Dim Str As String                                                      '3. 
    Dim Wynik As String                                                    '4. 
    With Sheets("Inicjały")                                                '5. 
        Do                                                                 '6. 
            WierszNr = WierszNr + 1                                        '7. 
            Str = ""                                                       '8. 
            Str = .Cells(WierszNr, 1).Value                                '9. 
            Wynik = ""                                                     '10. 
            For I = 1 To Len(Str)                                          '11. 
                If ((Mid(Str, I, 1) = " ") Or (Mid(Str, I, 1) = "-")) Then '12. 
                    Wynik = Wynik + Mid(Str, I + 1, 1)                     '13. 
                End If                                                     '12. 
            Next I                                                         '11. 
            .Cells(WierszNr, 2).Value = UCase(Mid(Str, 1, 1) + Wynik)      '14. 
        Loop Until (.Cells(WierszNr, 1).Value = "")                        '6. 
    End With                                                               '5. 
End Sub 
' 
'Legenda: 
'  1) Zadeklarowanie zmiennej "WierszNr", która jest typu liczbowego całkowitego. 
'     Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'     z zakresu od -2 147 483 648 do 2 147 483 647. 
'  2) Zadeklarowanie zmiennej "I", która jest typu liczbowego całkowitego. 
'     Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'     z zakresu od -2 147 483 648 do 2 147 483 647. 
'  3) Zadeklarowanie zmiennej "Str", która jest typu tekstowego. 
'     Zmienna tego typu umożliwia przechowywanie znaków alfanumerycznych. 
'  4) Zadeklarowanie zmiennej "Wynik", która jest typu tekstowego. 
'     Zmienna tego typu umożliwia przechowywanie znaków alfanumerycznych. 
'  5) Instrukcja wiążąca WITH ... END WITH umożliwia używanie innych instrukcji 
'     bez potrzeby pisania pełniej składni. Pod warunkiem, że te instrukcje znajdują 
'     się pomiędzy WITH a END WITH. 
'     Na przykład: 
'     Dzięki instrukcji wiążącej wystarczy napisać ".Cells(WierszNr, 1).Value" 
'     zamiast "Sheets("Inicjały").Cells(WierszNr, 1).Value". 
'  6) Pętla DO ... LOOP UNTIL(warunek). 
'     Pętla ta wykonuje instrukcje w niej zawarte określoną ilość razy. 
'     Ilość powtórzeń instrukcji w pętli zależy od warunku, który umieszczony jest na końcu 
'     pętli (za instrukcją UNTIL). Umieszczenie warunku na końcu pętli powoduje, że instrukcje 
'     w niej zawarte będą wykonane zawsze raz. 
'  7) WierszNr = WierszNr + 1. 
'     Jest to zwiększenie (tzw. inkrementacja) zawartości zmiennej liczbowej całkowitej 
'     "WierszNr" o wartość 1. 
'  8) Wyczyszczenie zmiennej tekstowej "Str". 
'  9) Przepisanie do zmiennej tekstowej "Str" zawartości komórki z kolumny nr 1 (w arkuszu jest to kolumna A) 
'     i z wiersza o nr, który przechowywany jest zmiennej liczbowej całkowitej "WierszNr". 
' 10) Wyczyszczenie zmiennej tekstowej "Wynik". 
' 11) Pętla FOR ... NEXT. 
'     Pętla wykonuje instrukcje w niej zawarte określoną ilość razy. Ilość powtórzeń jest określona 
'     za pomocą konstrukcji Len(Str). 
'     Len(P) - Instrukcja zwraca liczbę całkowitą określającą z ilu znaków składa się podany tekst. 
'              W miejsce parametru P umieszcza się zmienną tekstową przechowującą jakiś ciąg znaków. 
' 12) IF (warunek) THEN. 
'     Instrukcja warunkowa, której zadaniem jest sprawdzenie poprawności podanego warunku lub warunków. 
'     Warunek 1 - szukanie spacji. Warunek 2 - szukanie znaku "-". 
'     Wyszukanie obydwóch znaków odbywa się w podanym tekście. 
'     Jeżeli jeden z dwóch warunków będzie spełniony, to wykonaj instrukcje po słowie THEN. 
'     MID(Tekst, Start, Ile) - Wyciąga fragment tekstu z podanego ciągu znaków. 
'                              W miejsce parametru "Tekst" umieszcza się zmienną tekstową przechowującą dowolny tekst. 
'                              W miejsce parametru "Start" umieszcza się wartość liczbową całkowitą określającą początek pobierania fragmentu tekstu. 
'                              W miejsce parametru "Ile" umieszcza się wartość liczbową całkowitą określającą ile znaków ma zostać pobranych. 
' 13) Wynik = Wynik + Mid(Str, I + 1, 1). 
'     Dodanie do zmiennej tekstowej "Wynik" fragmentu tekstu, który został pobrany za pomocą instrukcji MID. 
'     Instrukcja ta wyciąga znak, który znajduje się na pozycji o jeden element dalej, niż znaki szukane 
'     w instrukcji warunkowej IF. 
' 14) Umieszczenie inicjału w komórce w kolumnie nr 2 (w arkuszu jest to kolumna B) w wierszu o nr, który przechowywany 
'     jest w zmiennej liczbowej całkowitej "WierszNr". Inicjały są pobrane za pomocą instrukcji MID 
'     i zmienione na wielkie litery za pomocą instrukcji UCASE. 