Sub ComboBoxSprDane(NazwaKomponentu As ComboBox, WartoscDomyslna As Boolean) 
'ComboBoxSprDane - Sprawdź, czy wybrana (lub wpisana) informacja jest na liście. 
'Copyright (c)by Jan T. Biernat 
'==== 
' 
'Wywołanie procedury: Call ComboBoxSprDane(nazwa_komponentu, WartoscDomyslna) 
'                     Drugi parametr (tj. "WartoscDomyslna") umożliwia ustawienie 
'                     sposobu działania procedury "ComboBoxSprDane", 
'                     w przypadku błędnego wpisu. Ustawienie wartości 2 parametru 
'                     na true, spowoduje pobranie 1 elementu listy jako wartości domyślnej. 
'                     Natomiast ustawienie wartości 2 parametru na false, spowoduje 
'                     podświetlenie komponentu na czerwono. 
' 
    Dim I As Integer 
    Dim CzyIstnieje As Boolean 
    CzyIstnieje = False 
    With NazwaKomponentu 
        .BackColor = &HFFFFFF 
        For I = 0 To .ListCount - 1 
            If (LCase(.Text) = LCase(.List(I))) Then 
                CzyIstnieje = True 
                Exit For 
            End If 
        Next I 
        If (CzyIstnieje = False) Then 
            'Wykonaj poniższe instrukcje, gdy podanego elementu nie ma na liście. 
            If (WartoscDomyslna = True) Then 
                .Text = .List(0) 
            Else 
                .BackColor = &HFF& 
            End If 
        End If 
    End With 
End Sub 