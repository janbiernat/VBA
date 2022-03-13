Private Sub TextBox1_Change() 
'TextBox1_Change() - Zaznacz odszukany element na liście. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim I As Integer 
    I = 0 
    With ListBox1 
        .MultiSelect = 0 
        For I = 0 To .ListCount - 1 
            .Selected(I) = False 'Odznacz wszystkie elementy na liście. 
        Next I 
        'Oznacz wyszukany element na liście. 
        If (Trim(TextBox1.Text) <> "") Then 
            For I = 0 To .ListCount - 1 
                If (LCase(Mid(Trim(.List(I, 0)), 1, Len(Trim(TextBox1.Text)))) = LCase(Trim(TextBox1.Text))) Then 
                    .Selected(I) = True 
                    Exit For 
                End If 
            Next I 
        End If 
    End With 
End Sub 