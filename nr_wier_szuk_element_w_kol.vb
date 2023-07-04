Function NrWierSzukElementK(ArkuszNazwa As String, Szukaj As String, NrKol As Integer, KolM As Integer) As Integer
'NrWierSzukElementK - Funkcja zwraca nr wiersza, w którym znajduje się szukany element.
'Copyright (c)by Jan T. Biernat
'=
'Wywałanie funkcji: Cells(3, 4).Value = NrWierSzukElementK("Informacje o uczniach", "Uczniowie zwolnieni z zajęć - cyfrowe technologie multimedialne w reklamie", 2, 999)
'
    Dim NrWier As Integer
    Dim I As Integer
    NrWier = 0
    I = 0
    ArkuszNazwa = Trim(ArkuszNazwa)
    Szukaj = Trim(Szukaj)
    If ((ArkuszNazwa <> "") And (Szukaj <> "")) Then
        If (NrKol < 1) Then
            NrKol = 1
        End If
        If (KolM < 1) Then
            KolM = 1
        End If
        With Sheets(ArkuszNazwa)
            For I = 1 To KolM
                If (LCase(Trim(.Cells(I, NrKol).Value)) = LCase(Szukaj)) Then
                    NrWier = I
                    Exit For
                End If
            Next I
        End With
        NrWierSzukElementK = NrWier
    Else
        NrWierSzukElementK = "BŁĄD -?Parametry funkcji są błędne!"
    End If
End Function