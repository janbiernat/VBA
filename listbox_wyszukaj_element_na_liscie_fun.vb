Sub ListBoxSzukaj(ListBoxNazwa As MSForms.ListBox, Szukaj As String) 
'ListBoxSzukaj - Funkcja zaznacza odnaleziony element na liście. 
'Copyright (c)by Jan T. Biernat 
' 
'Wywołanie procedury: Call ListBoxSzukaj(ListBox1, TextBox1.Text) 
' 
    Dim I As Integer 
    I = 0 
    Szukaj = Trim(Szukaj) 
    With ListBoxNazwa 
        .MultiSelect = 0 
        For I = 0 To .ListCount - 1 
            .Selected(I) = False 'Odznacz wszystkie elementy na liście. 
        Next I 
        'Oznacz wyszukany element na liście. 
        If (Szukaj <> "") Then 
            For I = 0 To .ListCount - 1 
                If (LCase(Mid(Trim(.List(I, 0)), 1, Len(Szukaj))) = LCase(Szukaj)) Then 
                    .Selected(I) = True 
                    Exit For 
                End If 
            Next I 
        End If 
    End With 
End Sub 