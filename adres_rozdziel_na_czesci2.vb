Sub adres_rozdziel_na_czesci2()
'Adres rozdziel na części 2 (c)by Jan T. Biernat
'=
'Napisz program, który rozdzieli podanych adres na części.
'Np.:
'    Podany jest adres: „Akademii Umiejętności 81/123”
'    lub adres: „Akademii Umiejętności 81m.123”.
'    W wyniku działania programu, otrzymujemy podzielony
'    adres na części, tj.:
'    1) Ulicę „Akademii Umiejętności”.
'    2) Nr bloku/mieszkania: 81.
'    3) Nr mieszkania: 123.
'=
'
    'Deklaracje stałych.
    Const Nag As Integer = 0 'Nagłówek: 0 - brak, 1 - jest.
    'Deklaracje zmiennych.
    Dim Wiersz As Integer
    Dim Str As String
    Dim I As Integer
    Dim AdrS As Integer
    Dim AdrU As Integer
    Dim AdrM As Integer
    'Wartości startowe.
    Wiersz = 0
    'Adres rozdziel na części.
    With Sheets("Arkusz1")
        Do
            Wiersz = Wiersz + 1
            Str = ""
            Str = Trim(Cells(Wiersz + Nag, 1).Value)
            If (Str <> "") Then
                I = 0
                AdrS = 0
                AdrU = 0
                AdrM = 0
                For I = Len(Str) To 1 Step -1
                    'Szukaj znaku ukośnika (tj. "/").
                    If ((Mid(Str, I, 1) = "/") And (AdrU = 0)) Then
                        AdrU = I
                    End If
                    'Szukaj skrótu (tj. "m.").
                    If ((LCase(Mid(Str, I, 2)) = "m.") And (AdrM = 0)) Then
                        AdrM = I
                    End If
                    'Szukaj znaku spacji (tj. " ").
                    If ((Mid(Str, I, 1) = " ") And (AdrS = 0)) Then
                        AdrS = I
                        Exit For
                    End If
                Next I
                'Wyczyść komórki w kolumnie nr 2, 3 i 4 (tj. kolumna B, C i D).
                Cells(Wiersz + Nag, 2).Value = "" 'Ulica.
                Cells(Wiersz + Nag, 3).Value = "" 'Nr bloku/domu.
                Cells(Wiersz + Nag, 4).Value = "" 'Nr mieszkania.
                'Ulica.
                If (AdrS > 0) Then
                    Cells(Wiersz + Nag, 2).Value = Trim(Mid(Str, 1, AdrS))                      'Ulica.
                    Cells(Wiersz + Nag, 3).Value = "'" + Trim(Mid(Str, AdrS, Len(Str)))         'Nr bloku/domu.
                Else
                    Cells(Wiersz + Nag, 2).Value = Str                                          'Ulica.
                End If
                'Nr bloku/domu i/lub nr mieszkania (Znak ukośnika "/").
                If ((AdrU > 0) And (AdrS > 0) And (AdrM = 0)) Then
                    Cells(Wiersz + Nag, 3).Value = "'" + Trim(Mid(Str, AdrS, (AdrU - AdrS)))    'Nr bloku/domu.
                    Cells(Wiersz + Nag, 4).Value = "'" + Trim(Mid(Str, AdrU + 1, Len(Str)))     'Nr mieszkania.
                Else
                    'Nr bloku/domu i/lub nr mieszkania (Skrót "m.").
                    If ((AdrU = 0) And (AdrS > 0) And (AdrM > 0)) Then
                        Cells(Wiersz + Nag, 3).Value = "'" + Trim(Mid(Str, AdrS, (AdrM - AdrS)))    'Nr bloku/domu.
                        Cells(Wiersz + Nag, 4).Value = "'" + Trim(Mid(Str, AdrM + 2, Len(Str)))     'Nr mieszkania.
                    End If
                End If
            End If
        Loop Until (.Cells(Wiersz + Nag, 1).Value = "")
    End With
End Sub