Sub ListBoxOdznaczElementy(ListBoxNazwa As MSForms.ListBox)
'ListBoxOdznaczElementy - Funkcja odznacza wszystkie elementy na liście.
'Copyright (c)by Jan T. Biernat
'
'Wywołanie procedury: Call ListBoxOdznaczElementy(ListBox1)
'
    Dim I As Integer
    I = 0
    With ListBoxNazwa
        .MultiSelect = 0                    'MultiSelectSingle.
        For I = .ListCount - 1 To 0 Step -1
            .Selected(I) = False            'Odznacz wszystkie elementy na liście.
        Next I
    End With
End Sub