Sub tabliczka_mnozenia()
'Tabliczka mnożenia na krzyż do 100 (10 * 10 = 100).
'
    Dim A As Integer                            '1.
    Dim B As Integer                            '2.
    For A = 1 To 10                             '3.
        Cells(1, 1 + A).Value = A               '4.
        For B = 1 To 10                         '5.
            Cells(1 + B, 1).Value = B           '6.
            Cells(1 + B, 1 + A).Value = (A * B) '7.
        Next B                                  '5.
    Next A                                      '3.
    Range("A1").Select                          '8.
End Sub
'
'Legenda:
'1. Zadeklarowanie zmiennej o nazwie "A", która jest
'   typu liczbowego całkowitego. Zmienne tego typu zajmują
'   4 bajty i są wstanie przechowywać liczby z zakresu
'   od -32 768 do 32 768.
'2. Zadeklarowanie zmiennej o nazwie "B", która jest
'   typu liczbowego całkowitego. Zmienne tego typu zajmują
'   4 bajty i są wstanie przechowywać liczby z zakresu
'   od -32 768 do 32 768.
'3. Pętla FOR ... NEXT.
'   Pętla FOR ... NEXT oraz instrukcje zawarte pomiędzy FOR a NEXT będą powtórzone/wykonane 10 razy.
'   Pomiędzy FOR a TO umieszcza się zmienną z przypisaną wartością startową (np. A = 1).
'   Natomiast po słowie TO wpisujemy wartość określającą ile razy pętla ma być powtórzona/wykonana (np. 10).
'4. Wpisanie do komórek wartości liczbowych przechowywanych w zmiennej liczbowej "A",
'   rozpoczynając wpisywanie od kolumny nr 2 (tj. od kolumny B) i kończąc na kolumnie nr 11 (tj. kolumnie K) w wierszu nr 1.
'   Zawartość zmiennej liczbowej "A" jest ustawiana przez pętlę przy każdym powtórzeniu.
'5. Pętla FOR ... NEXT.
'   Pętla FOR ... NEXT oraz instrukcje zawarte pomiędzy FOR a NEXT będą powtórzone/wykonane 10 razy.
'   Pomiędzy FOR a TO umieszcza się zmienną z przypisaną wartością startową (np. B = 1).
'   Natomiast po słowie TO wpisujemy wartość określającą ile razy pętla ma być powtórzona/wykonana (np. 10).
'6. Wpisanie do komórek wartości liczbowych przechowywanych w zmiennej liczbowej "B"
'   w kolumnie nr 1 (tj kolumnie A), rozpoczynając wpisywanie od wiersza nr 2 i kończąc na wierszu nr 11.
'7. Wpisanie wyniku mnożenia dwóch liczb w poszczególnych komórkach,
'   rozpoczynając od komórki w wierszu nr 2 i kolumny nr 2 (tj. od komórki B2).
'   W trakcie powtórzeń poszczególnych pętli FOR nr kolumny i nr wiersza zmienia się
'   przez zmianę zawartości zmiennych "A" i "B". Zmianę zawartości zmiennych generują
'   poszczególne pętle FOR.
'   W celu umożliwienia przemnożenia na krzyż nagłówku kolumny i wiersza,
'   trzeba ustawić początek na komórce B2. Jest to możliwe tylko wtedy, gdy
'   do zmiennych "A" i "B" dodamy wartość 1.
'8. Zaznaczenie komórki A1, która znajduje się w wierszu nr 1 i w kolumnie nr 1.