Function SzukajDanych(ArkuszNazwa As String, Szukaj As String, CzyDanePoPrawej As Boolean, WierszM As Integer, KolM As Integer) As String
'SzukajDanych - Funkcja znajduje dane na podstawie opisu.
'Copyright (c)by Jan T. Biernat
'=
'Wywołanie funkcji: Cells(2, 3).Value = SzukajDanych("Arkusz2", "nieodpowiednie", False, 19, 20)
'
    Dim W As Integer
    Dim K As Integer
    Dim Str As String
    W = 0
    K = 0
    Str = ""
    ArkuszNazwa = Trim(ArkuszNazwa)
    Szukaj = Trim(Szukaj)
    If ((ArkuszNazwa <> "") And (Szukaj <> "")) Then
        If (WierszM < 1) Then
            WierszM = 1
        End If
        If (KolM < 1) Then
            KolM = 1
        End If
        With Sheets(ArkuszNazwa)
            For K = 1 To KolM
                For W = 1 To WierszM
                    If (LCase(Trim(.Cells(W, K).Value)) = LCase(Szukaj)) Then
                        If (CzyDanePoPrawej = True) Then
                            Str = Trim(.Cells(W, K + 1).Value) 'Dane po prawej stronie opisu w tym samym wierszu.
                        Else
                            Str = Trim(.Cells(W + 1, K).Value) 'Dane po poniżej opisu w tej samej kolumnie.
                        End If
                        Exit For
                    End If
                Next W
            Next K
        End With
        SzukajDanych = Str
    Else
        SzukajDanych = "BŁĄD -?Pierwsze dwa parametry są puste!"
    End If
End Function