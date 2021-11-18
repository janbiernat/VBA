Sub tabliczka_mnozenia() 
'Tabliczka mno�enia na krzy� do 100 (10 * 10 = 100). 
' 
    Dim A As Integer                          '1. 
    Dim B As Integer                          '2. 
    For A = 1 To 10                           '3. 
        Cells(3, 1 + A).Value = A             '4. 
        For B = 1 To 10                       '5. 
            Cells(3 + B, 1).Value = B         '6. 
            Cells(3 + B, 1 + A).Value = A * B '7. 
        Next B                                '5. 
    Next A                                    '3. 
End Sub 
' 
'Legenda: 
' 1) Deklaracja zmiennej liczbowej ca�kowitej "A". 
' 2) Deklaracja zmiennej liczbowej ca�kowitej "B". 
' 3) P�tla FOR ... NEXT. 
'    P�tla FOR ... NEXT oraz instrukcje zawarte pomi�dzy FOR a NEXT b�d� wykonane 10 razy. 
'    Pomi�dzy FOR a TO umieszcza si� zmienn� z przypisan� warto�ci� startow� (np. A = 1). 
'    Natomiast po s�owie TO wpisujemy warto�� okre�laj�c� ile razy p�tla ma by� wykonana (np. 10). 
' 4) Wpisanie do kom�rki warto�ci liczbowej przechowywanej w zmiennej liczbowej "A". 
'    Dotyczy to kom�rek, kt�re znajduj� si� w wierszu nr 3 w kolejnych kolumnach zaczynaj�c 
'    od kolumny nr 2 (tj. 1+A). 
' 5) P�tla FOR ... NEXT. 
'    P�tla FOR ... NEXT oraz instrukcje zawarte pomi�dzy FOR a NEXT b�d� wykonane 10 razy. 
'    Pomi�dzy FOR a TO umieszcza si� zmienn� z przypisan� warto�ci� startow� (np. B = 1). 
'    Natomiast po s�owie TO wpisujemy warto�� okre�laj�c� ile razy p�tla ma by� wykonana (np. 10). 
' 6) Wpisanie do kom�rki warto�ci liczbowej przechowywanej w zmiennej liczbowej "B". 
'    Dotyczy to kom�rek, kt�re znajduj� si� w kolumnie nr 1 w kolejnych wierszach. 
'    Zaczynaj�c od wiersza nr 4 (tj. 3+B). 
' 7) Wpisanie wyniku mno�enia dw�ch liczb w poszczeg�lnych kom�rkach. 
'    Zaczynaj�c od kom�rki, kt�ra jest w wierszu nr 4 i w kolumnie nr 2. 