Sub nr_wierszy() 
'Numerowanie wierszy. 
' 
    Dim Licznik As Integer                '1 
    Licznik = 0                           '2 
    Do                                    '3 
        Licznik = Licznik + 1             '4 
        Cells(Licznik, 1).Value = Licznik '5 
    Loop Until (Licznik > 20)             '6 
End Sub 
' 
'Legenda: 
'1) Zadeklarowanie zmiennej liczbowej całkowitej "Licznik". 
'2) Przypisanie wartości zerowej do zmiennej liczbowej całkowitej "Licznik" (tj. Licznik = 0). 
'3) Początek pętli DO ... LOOP UNTIL(warunek). 
'   Instrukcje zawarte pomiędzy DO a LOOP UNTIL będą 
'   wykonywane tak długo, jak długo jest spełniony umieszczony na końcu warunek. 
'   Umieszczenie warunku na końcu powoduje, że instrukcje 
'   Instrukcje wewnątrz pętli będą zawsze wykonane minimum raz. 
'   Nawet, jeżeli warunek postawiony na końcu nie będzie spełniony. 
'   Spowodowane jest to koniecznością przejścia wszystkich instrukcji 
'   zawartych wewnątrz pętli zanim zostanie sprawdzony warunek. 
'4) Zwiększenie zawartości zmiennej liczbowej całkowitej "Licznik" o wartość 1. 
'   Na początku do zmiennej "Licznik" jest przypisana wartość 0 
'   (patrz linia nad pętlą DO ... LOOP UNTIL - komentarz nr 2). 
'   Przy każdym wykonaniu instrukcji znajdujących się wewnątrz pętli DO ... LOOP UNTIL 
'   zawartość zmiennej "Licznik" jest zwiększana o wartość 1. 
'   Czyli do początkowej zawartości zmiennej "Licznik" (tj. wartość 0) jest 
'   dodawana wartość 1 i od tej pory zmienna "Licznik" przechowuje wartość 1. 
'   Przy drugim wykonaniu instrukcji znajdujących się wewnątrz pętli ponownie 
'   dodawana jest wartość 1 i od tej pory zmienna "Licznik" przechowuje wartość 2. 
'   Zwiększanie zawartości zmiennej "Licznik" będzie tak długo wykonywane, jak 
'   długo jest spełniony umieszczony na końcu warunek. 
'5) Wyświetlenie zawartości zmiennej liczbowej całkowitej "Licznik" 
'   w kolumnie 1 w kolejnych wierszach licząc od wiersza 1. 
'   Nr wiersza przechowywany jest w zmiennej liczbowej całkowitej "Licznik". 
'6) Sprawdzanie warunku, czy zawartość zmiennej "Licznik" jest większa do wartości 20. 
'   Jeżeli tak, to przerwij wykonanie pętli DO ... LOOP UNTIL. 