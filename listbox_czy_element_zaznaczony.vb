Function ListBoxCzyElementZaznaczony(ListBoxNazwa As MSForms.ListBox) As Boolean
'ListBoxCzyElementZaznaczony - Funkcja sprawdza, czy na liście został zaznaczony element.
'Copyright (c)by Jan T. Biernat
'
'Wywołanie procedury: If ((ListBoxCzyElementZaznaczony(ListBox1) = True) And ( ...
'
    Dim I As Integer
    Dim Zaznaczony As Boolean
    I = 0
    Zaznaczony = False
    With ListBoxNazwa
        For I = .ListCount - 1 To 0 Step -1
            If (.Selected(I) = True) Then
                Zaznaczony = True
                Exit For
            End If
        Next I
    End With
    ListBoxCzyElementZaznaczony = Zaznaczony
End Function