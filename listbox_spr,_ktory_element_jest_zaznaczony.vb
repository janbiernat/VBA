Private Sub CommandButton1_Click() 
'ListBox: Sprawdzenie, który element jest zaznaczony. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim I As Integer                                              '1. 
    Dim W As Integer                                              '2. 
    W = 0                                                         '3. 
    With ListBox1                                                 '4. 
        For I = 0 To .ListCount - 1                               '5. 
            If (.Selected(I) = True) Then                         '6. 
                W = W + 1                                         '7. 
                Sheets("Arkusz1").Cells(W, 1).Value = .List(I, 0) '8. 
            End If                                                '6. 
        Next I                                                    '5. 
    End With                                                      '4. 
End Sub 
' 
'Legenda: 
'Powyższy kod zadziała, jeżeli będą ustawione następujące właściwości 
'komponentu ListBox1: 
' > Właściwość ListStyle na 1 - fmListStyleOption. 
' > Właściwość MultiSelect na 1 - fmMultiSelectMulti. 
' 
' 1) Zadeklarowanie zmiennej liczbowej całkowitej "I". 
'    Typ danych typu integer zajmuje 2 bajty pamięci 
'    i może przyjmować liczby z zakresu od -32 768 do 32 767. 
' 2) Zadeklarowanie zmiennej liczbowej całkowitej "W". 
' 3) Przypisanie do zmiennej liczbowej całkowitej "W" wartości równiej 0. 
' 4) Zastosowanie konstrukcji wiążącej WITH ... END WITH. 
'    Po słowie WITH umieszcza się nazwę np. komponentu (np. ListBox1). 
'    Konstrukcja ta umożliwia skrócenie pisanego kodu o np. nazwę 
'    komponentu. Zamiast pisać "ListBox1.ListCount - 1", to dzięki 
'    konstrukcji wiążącej, kod skrócony jest do ".ListCount - 1". 
'    Dotyczy to wszystkich właściwości komponentu ListBox1. 
' 5) Pętla FOR. 
'    Pętla wykona znajdujący się w niej blok instrukcji, tyle razy ile jest 
'    elementów na liście. 
'    Ilość elementów na liście jest określona za pomocą instrukcji .ListCount - 1. 
' 6) Instrukcja warunkowa (tj. IF (warunek) THEN). 
'    Sprawdzenie jakie elementy na liście są zaznaczone. 
'    Jeżeli warunek będzie spełniony (tj. element jest zaznaczony), 
'    to wykonaj instrukcje po słowie THEN. 
' 7) W = W+1. 
'    Jest to tak zwana inkrementacja, czyli zwiększenie zawartości zmiennej 
'    liczbowej "W" i wartość 1. 
'    Przy każdym przebiegu pętli FOR zawartość zmiennej liczbowej "W" jest 
'    zwiększana o wartość 1. Przy 1 przebiegu zmienna liczbowa "W" przechowuje 
'    wartość 1. Przy 2 przebiegu zmienna liczbowa "W" przechowuje wartość 2. 
'    I tak do wykonania ostatniego przebiegu pętli FOR. 
'    UWAGA: Zawartość zmiennej liczbowej "W" będzie zwiększana o wartość 1, gdy 
'    zostanie spełniony warunek (tj. "If (.Selected(I) = True) Then") pod 
'    nr komentarza 6. 
' 8) Pobranie elementu (tj. tekstu) z listy, z pozycji o numerze przechowywanym 
'    w zmiennej liczbowej "I". Następnie przypisanie pobranego tekstu do instrukcji 
'    CELLS. Instrukcja CELLS umożliwia umieszczanie danych w komórkach arkusza, 
'    jak i pobieranie danych z komórek arkusza. 
'    Za pomocą instrukcji SHEETS, określany jest arkusz w którym dane będą umieszczane. 
' 
