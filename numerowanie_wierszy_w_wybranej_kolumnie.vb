Sub WierszeNrP(ArkuszNazwa As String, WierszNrOd As Long, KolumnaNr As Long)
'WierszeNrP - Numerowanie wierszy w wybranej kolumnie.
'Copyright (c)by Jan T. Biernat
'
'Wywołanie procedury:
'Call WierszeNumerowanie(zmienna_tekstowa, zmienna_liczbowa)
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