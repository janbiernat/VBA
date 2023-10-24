Sub wiersz_usun_duplikaty()
'WIERSZ: Usuń duplikaty.
'Copyright (c)by Jan T. Biernat
'
    Dim Licznik1 As Integer
    Dim Licznik2 As Integer
    Dim I As Integer
    Dim Jest As Boolean
    Licznik1 = 0
    I = 0
    Do
        Licznik2 = 0
        Jest = False
        Licznik1 = Licznik1 + 1
        '
        'Sprawdź, czy w kolumnie B jest dodawany tekst/liczba.
        Do
            Licznik2 = Licznik2 + 1
            If (LCase(Trim(Cells(Licznik1, 1).Value)) = LCase(Trim(Cells(Licznik2, 2).Value))) Then
                Jest = True
                Exit Do
            End If
        Loop Until (Trim(Cells(Licznik2, 2).Value) = "")
        '
        'Jeżeli tekstu/liczby nie ma w kolumnie B, to dodaj tekst/liczbę do kolejnej komórki w kolumnie B.
        If (Jest = False) Then
            Cells(Licznik2, 2).Value = Trim(Cells(Licznik1, 1).Value)
        End If
    Loop Until (Trim(Cells(Licznik1, 1).Value) = "")
    'Kopiowanie tekstu/liczby z kolumny B do kolumny A.
    For I = 1 To Licznik1
        Cells(I, 1).Value = Trim(Cells(I, 2).Value)
        Cells(I, 2).Value = ""
    Next I
End Sub