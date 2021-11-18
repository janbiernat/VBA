Private Sub Workbook_Open() 
'Tabliczka mnożenia - Losowanie 1 kolumny z 10 kolumn. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim A As Byte                                     '1. 
    Dim I As Byte 
    Randomize                                         '2. 
    A = 0                                             '3. 
    A = Int((10 * Rnd) + 1)                           '4. 
    'Wyświetl wylosowaną kolumnę na ekranie. 
    With Sheets("Arkusz1")                            '5. 
        .Cells(1, 1).Value = "Kolumna nr: " + CStr(A) '6. 
        For I = 1 To 10                               '7. 
            Cells(2 + I, 1).Value = I                 '8. 
            Cells(2 + I, 2).Value = " * " 
            Cells(2 + I, 3).Value = A 
            Cells(2 + I, 4).Value = " = " 
            Cells(2 + I, 5).Value = (I * A)           '9. 
        Next I                                        '7. 
    End With                                          '5. 
    ' 
    'Legenda: 
    '1. Deklaracja zmiennej liczbowej całkowitej "A". Zmienna jest typu Byte. 
    '   Przechowywać może liczby z zakresu o 0 do 255 i zajmuje 1B pamięci komputera. 
    '2. Zainicjowanie generatora liczb losowych. 
    '3. Przypisanie wartości zerowej do zmiennej liczbowej "A". 
    '4. Generuje losową liczbę z zakresu od 1 do 10 i przypisuje zmiennej liczbowej "A". 
    '   Rnd - instrukcja zwraca liczbę losową z przedziału od 0 do 1. 
    '         Instrukcja jest podobna do funkcji LOS w arkuszu Excel. 
    '   Int - instrukcja służy do zaokrąglania liczb posiadających miejsca dziesiętne do liczb całkowitych. 
    '         Na przykład: Int(3.14) instrukcja zwróci wartość całkowitą 3. 
    '5. Konstrukcja wiążąca WITH ... END WITH. 
    '   Po słowie With umieszcza się np. nazwę arkusza [tj. Sheets("Arkusz1")] 
    '   i od tego momentu nie trzeba poprzedzać żadnej instrukcji tą nazwą arkusza. 
    '   Gdyby w tym przykładzie nie było instrukcji wiążącej WITH ... END WITH, to 
    '   przed każdą instrukcją Cells trzeba byłoby umieszczać konstrukcję 
    '   Sheets("Arkusz1").Cells(1, 1).Value = ... . 
    '6. Umieszczenie w komórce A1 (wiersz 1 i kolumna 1 w arkuszu) tekstu znajdującego 
    '   pomiędzy cudzysłowami (tj. "Kolumna nr: ") oraz zamienionej(przekonwertowanej) na ciąg znaków zawartości zmiennej liczbowej "A". 
    '   CStr(zmienna_liczbowa) - Konwertuje typ liczbowy na ciąg znaków. 
    '                            Na przykład: Liczba 123 to ciąg znaków składający się 
    '                                         z trzech znaków "1", "2" i "3". 
    '7. Pętla FOR, która będzie powtarzana 10 razy. 
    '8. Umieszczenie zawartości zmiennej liczbowej "I" do kolejnych komórek (rozpoczynając od komórki A3 w arkuszu). 
    '   Wszystko odbywa się w kolumnie 1 arkusza. 
    '9. Umieszczenie w kolejnych komórkach (rozpoczynając od komórki E3 w arkuszu w kolumnie nr 5) 
    '   wyniku mnożenia dwóch liczb. Liczby te są przechowywane w zmiennych liczbowych "I" i "A". 
    '   W zmiennej liczbowej "I" przechowywana jest wartość nadawana przez pętlę FOR od liczby 1 do 10. 
    '   Natomiast w zmiennej Liczbowej "A" jest przechowywana wartość, która została wylosowana za pomocą instrukcja Int. 
End Sub 