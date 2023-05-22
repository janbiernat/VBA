Sub KopiujDaneOdKonca() 
'KopiujDaneOdKonca (c)by Jan T. Biernat 
'Funkcja kopiuje dane z kolumny A do kolumny B w odwrotnej kolejności. 
' 
    Dim I As Integer                                    '1. 
    Dim W As Integer                                    '2. 
    I = 0                                               '3. 
    W = 0                                               '4. 
    With Sheets("Arkusz1")                              '5. 
        Do                                              '6. 
            W = W + 1                                   '7. 
        Loop Until (.Cells(W, 1).Value = "")            '6. 
        For I = 1 To (W - 1)                            '8. 
            Cells(W - I, 2).Value = Cells(I, 1).Value   '9. 
        Next I                                          '8. 
    End With                                            '5. 
End Sub 
' 
'Legenda: 
'1. Zadeklarowanie zmiennej liczbowej o nazwie "I", która jest typu liczbowego całkowitego. 
'   Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'   z zakresu od -2 147 483 648 do 2 147 483 647. 
'2. Zadeklarowanie zmiennej liczbowej o nazwie "W", która jest typu liczbowego całkowitego. 
'   Zmienne tego typu zajmują 4 bajty i są wstanie przechowywać liczby 
'   z zakresu od -2 147 483 648 do 2 147 483 647. 
'3. Przypisanie wartości zerowej do zmiennej liczbowej całkowitej "I". 
'4. Przypisanie wartości zerowej do zmiennej liczbowej całkowitej "W". 
'5. With Sheets("Arkusz1") 
'   Instrukcja wiążąca WITH ... END WITH umożliwia używanie innych instrukcji 
'   bez potrzeby pisania pełniej składni. Pod warunkiem, że te instrukcje znajdują 
'   się pomiędzy WITH a END WITH. 
'   Na przykład: 
'                Dzięki instrukcji wiążącej wystarczy napisać ".Cells(W, 1).Value" 
'                zamiast "Sheets("Arkusz1").Cells(W, 1).Value". 
'6. Pętla DO ... LOOP UNTIL(warunek). 
'   Pętla ta wykonuje instrukcje w niej zawarte określoną ilość razy. 
'   Ilość powtórzeń instrukcji (lub bloku instrukcji) w pętli zależy od warunku, 
'   który umieszczony jest na końcu pętli (za instrukcją UNTIL). 
'   Umieszczenie warunku na końcu pętli powoduje, że instrukcja (lub blok instrukcji) 
'   zawsze będą wykonane raz. 
'   W tym przykładzie pętla DO ... LOOP UNTIL(warunek) zakończy swoje działanie w momencie 
'   natrafienia na pustą komórkę. 
'7. Zwiększenie zawartości zmiennej liczbowej całkowitej "W" o wartość 1. 
'8. Pętla FOR ... NEXT. 
'   Pętla wykonuje instrukcje w niej zawarte określoną ilość razy. 
'   O ilości powtórzeń decyduje zapis, który znajduje się za słowem TO 
'   (np. "(W - 1)" - Ilość powtórzeń jest przechowywana w formie liczby całkowitej 
'    w zmiennej "W" pomniejszonej o wartość 1).
'9. Przepisanie danych w odwrotnej kolejności z kolumny A (nr 1) z poszczególnych komórek 
'   do kolumny B (nr 2) do poszczególnych komórek.
'   Cells(I, 1).Value - pobranie danych z kolumny A (nr 1) z poszczególnych komórek, licząc 
'                       od 1 komórki i przepisanie ich (tych danych) do kolumny B (nr 2) 
'                       do poszczególnych komórek, licząc od końca. 
'   Cells(W - I, 2).Value - wpisanie danych pobranych z kolumny A (nr 1) z poszczególnych komórek 
'                           do kolumny B (nr 2) w poszczególnych komórkach. 