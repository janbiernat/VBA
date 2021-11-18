Sub tabliczka_mnozenia() 
'Tabliczka mno¿enia na krzy¿ do 100 (10 * 10 = 100). 
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
' 1) Deklaracja zmiennej liczbowej ca³kowitej "A". 
' 2) Deklaracja zmiennej liczbowej ca³kowitej "B". 
' 3) Pêtla FOR ... NEXT. 
'    Pêtla FOR ... NEXT oraz instrukcje zawarte pomiêdzy FOR a NEXT bêd¹ wykonane 10 razy. 
'    Pomiêdzy FOR a TO umieszcza siê zmienn¹ z przypisan¹ wartoœci¹ startow¹ (np. A = 1). 
'    Natomiast po s³owie TO wpisujemy wartoœæ okreœlaj¹c¹ ile razy pêtla ma byæ wykonana (np. 10). 
' 4) Wpisanie do komórki wartoœci liczbowej przechowywanej w zmiennej liczbowej "A". 
'    Dotyczy to komórek, które znajduj¹ siê w wierszu nr 3 w kolejnych kolumnach zaczynaj¹c 
'    od kolumny nr 2 (tj. 1+A). 
' 5) Pêtla FOR ... NEXT. 
'    Pêtla FOR ... NEXT oraz instrukcje zawarte pomiêdzy FOR a NEXT bêd¹ wykonane 10 razy. 
'    Pomiêdzy FOR a TO umieszcza siê zmienn¹ z przypisan¹ wartoœci¹ startow¹ (np. B = 1). 
'    Natomiast po s³owie TO wpisujemy wartoœæ okreœlaj¹c¹ ile razy pêtla ma byæ wykonana (np. 10). 
' 6) Wpisanie do komórki wartoœci liczbowej przechowywanej w zmiennej liczbowej "B". 
'    Dotyczy to komórek, które znajduj¹ siê w kolumnie nr 1 w kolejnych wierszach. 
'    Zaczynaj¹c od wiersza nr 4 (tj. 3+B). 
' 7) Wpisanie wyniku mno¿enia dwóch liczb w poszczególnych komórkach. 
'    Zaczynaj¹c od komórki, która jest w wierszu nr 4 i w kolumnie nr 2. 