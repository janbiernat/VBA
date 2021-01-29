Private Sub Workbook_Open() 
'Auostart. 
'Automatyczne numerowane wierszy podczas uruchomienia arkusza kalkulacyjnego MS Excel. 
' 
    Dim Licznik As Integer                      '1 
    Licznik = 0                                 '2 
    Do                                          '3 
        Licznik = Licznik + 1                   '4 
        If (Cells(Licznik, 2).Value <> "") Then '5 
            Cells(Licznik, 1).Value = Licznik   '6 
        End If                                  '5 
    Loop Until (Cells(Licznik, 2).Value = "")   '3 
' 
'Legenda: 
'1) Deklaracja zmiennej liczbowej całkowitej o nazwie "Licznik". 
'2) Przypisanie wartości 0 do zmiennej liczbowej "Licznik". 
'3) Początek pętli DO .. LOOP UNTIL(warunek). 
'   Pomiędzy DO a LOOP UNTIL umieszcza się instrukcje, 
'   które wykonywane są z góry określoną ilość razy. 
'   O ilości powtórzeń pętli decyduje warunek umieszczony 
'   za instrukcjami LOOP UNTIL. Umieszczenie warunku na końcu 
'   pętli powoduje, że instrukcje umieszczone pomiędzy DO .. LOOP UNTIL 
'   będą zawsze raz wykonane. Bez względu, czy warunek jest spełniony, czy nie. 
'   Pętla DO .. LOOP UNTIL zakończy swoje działanie w momencie natrafienia na pustą 
'   komórkę w kolumnie nr 2 (w arkuszu to kolumna B). 
'4) Zwiększenie zmiennej liczbowej całkowitej "Licznik" o wartość 1. 
'   Zmienna ta będzie zwiększana przy każdym powtórzeniu pętli DO .. LOOP UNTIL. 
'   Ta konstrukcja nosi nazwę inkrementacji. 
'5) Instrukcja warunkowa IF (warunek) THEN sprawdza, czy komórka w kolejnym 
'   wierszu w kolumnie nr 2 (w arkuszu to kolumna B) zawiera jakieś dane (np. tekst). 
'   Jeżeli tak, to wykonaj instrukcje znajdujące się pomiędzy THEN a END IF. 
'6) Cells(Licznik, 1).Value = Licznik 
'   Wpisanie do komórki w kolejnych wierszach w kolumnie nr 1 (w arkuszu to kolumna A) 
'   kolejnego numeru (numer przechowywany jest w zmiennej liczbowej "Licznik"). 
End Sub 