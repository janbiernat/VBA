
Sub WierszeNrP(ArkuszNazwa As String, WierszNrOd As Long, KolumnaNr As Long)
'WierszeNrP - Numerowanie wierszy w wybranej kolumnie.
'Copyright (c)by Jan T. Biernat
'
'Wywołanie procedury:
'Call WierszeNrP(nazwa_arkusza, nr_wiersza, nr_kolumny)
'
    Dim Licznik As Long
    Licznik = 0
    With Sheets(ArkuszNazwa)
        If ((KolumnaNr > 1) And (WierszNrOd > 0)) Then
            Do
                Licznik = Licznik + 1
                If (.Cells(WierszNrOd + Licznik, KolumnaNr).Value <> "") Then
                    If (.Cells(WierszNrOd + Licznik, KolumnaNr - 1).Value = "") Then
                        .Cells(WierszNrOd + Licznik, KolumnaNr - 1).Value = Licznik
                    End If
                End If
            Loop Until (.Cells(WierszNrOd + Licznik, KolumnaNr).Value = "")
        End If
    End With
End Sub

Sub numeruj_wiersze_procedura()
'Numerowanie wierszy w wybranej kolumnie.
'Copyright (c)by Jan Biernat
'
    'Ponumerowanie dni tygodnia z 1 kolumny
     Call WierszeNrP("Dane", 1, 1)
    'Ponumerowanie miesięcy z 2 kolumny
     Call WierszeNrP("Dane", 1, 2)
    'Ponumerowanie dni tygodnia z 4 kolumny
     Call WierszeNrP("Dane", 1, 4)
    'Ponumerowanie miesięcy z 6 kolumny
     Call WierszeNrP("Dane", 1, 6)
    'Ponumerowanie imion z 8 kolumny
     Call WierszeNrP("Dane", 4, 8)
End Sub