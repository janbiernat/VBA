Sub ComboBoxSprDane2(NazwaKomponentu As ComboBox)
'ComboBoxSprDane2 - Sprawdź, czy wpisana informacja jest na liście.
'Copyright (c)by Jan T. Biernat
'=
'
'Wywołanie procedury: Call ComboBoxSprDane2(ComboBox1)
'
    Dim I As Integer
    I = 0
    With NazwaKomponentu
        .BackColor = &HFF& 'Ustaw kolor czerwony dla tła.
        For I = 0 To .ListCount - 1
            If (LCase(Trim(.Text)) = LCase(Trim(.List(I)))) Then
                .BackColor = &HFFFFFF 'Ustaw kolor biały dla tła.
                Exit For
            End If
        Next I
    End With
End Sub