Function SprPoprOdp(WierszNr As Integer) As Integer
'SprPoprOdp - Sprawdź poprawność odpowiedzi.
'Copyright (c)by Jan T. Biernat
'=
'Moduł sprawdza poprawność odpowiedzi,
'porównując z odpowiedziami wzorcowymi.
'Odpowiedzi pochodzą z testu jednokrotnego wyboru.
'Jest to test w którym jedna odpowiedź jest prawidłowa
'z pośród czterech możliwych do wyboru odpowiedzi.
'
    Dim T1 As String
    Dim T2 As String
    Dim Jest As Integer
    Jest = 0
    'Zabezpieczenie nr wiersza.
        If (WierszNr < 1) Then
            WierszNr = 1
        End If
    'Zabezpieczenie kolumn nr 1 i 2 (tj. kolumn A i B).
        If ((Trim(Cells(WierszNr, 1).Value) <> "") And (Trim(Cells(WierszNr, 2).Value) <> "")) Then
            'Kolumna 1 (Kolumna A).
                T1 = ""
                T1 = UCase(Trim(Mid(Cells(WierszNr, 1).Value, 1, 1)))
                Cells(WierszNr, 1).Value = T1
            'Kolumna 2 (Kolumna B).
                T2 = ""
                T2 = UCase(Trim(Mid(Cells(WierszNr, 2).Value, 1, 1)))
                Cells(WierszNr, 2).Value = T2
            'Porównanie dwóch odpowiedzi (odpowiedzi wzorcowej z odpowiedzią z testu jednokrotnego wyboru).
                If (Cells(WierszNr, 1).Value = Cells(WierszNr, 2).Value) Then
                    SprPoprOdp = 1
                Else
                    SprPoprOdp = 0
                End If
        Else
            SprPoprOdp = -1
        End If
End Function

Sub TestSprOdp()
'TestSprOdp - Sprawdź odpowiedzi.
'Copyright (c)by Jan T. Biernat
'
    Dim WierszNr As Integer
    Dim Status As Integer
    Dim OdpDobre As Integer
    Dim OdpZle As Integer
    Dim Procent As Integer
    Dim RaportPokaz As Boolean
    Dim JakaOcena As String
    WierszNr = 0
    OdpDobre = 0
    OdpZle = 0
    RaportPokaz = False
    'Nagłówek.
        Cells(1, 1).Value = "Odpowiedzi wzorcowe"
        Cells(1, 2).Value = "Odpowiedzi ucznia"
        Cells(1, 3).Value = "Status"
        Cells(1, 4).Value = "Komentarz"
    'Sprawdź odpowiedzi na postawie odpowiedzi wzorcowych.
        Do
            WierszNr = WierszNr + 1
            Status = 0
            If ((Trim(Cells(WierszNr + 1, 1).Value) <> "") And (Trim(Cells(WierszNr + 1, 2).Value) <> "")) Then
                RaportPokaz = True
                Status = SprPoprOdp(WierszNr + 1)
                If (Status = -1) Then
                    'Gdy dowolna komórka w danym wierszu jest pusta w kolumnie 1(A) lub kolumnie 2(B).
                        Cells(WierszNr + 1, 3).Value = "?"
                        Cells(WierszNr + 1, 4).Value = "BŁĄD -?Komórka w kolumnie 1(A) lub 2(B) jest pusta!"
                ElseIf (Status > 0) Then
                        'Odpowiedzi poprawne.
                            Cells(WierszNr + 1, 3).Value = 1
                            Cells(WierszNr + 1, 4).Value = ""
                            OdpDobre = OdpDobre + 1
                    Else
                        'Odpowiedzi błędne.
                            Cells(WierszNr + 1, 3).Value = 0
                            Cells(WierszNr + 1, 4).Value = "<- Źle!"
                            OdpZle = OdpZle + 1
                    End If
            End If
        Loop Until (Trim(Cells(WierszNr + 1, 1).Value) = "")
    'Raport
        If (RaportPokaz = True) Then
            Cells(WierszNr + 2, 1).Value = "Raport"
            Cells(WierszNr + 3, 1).Value = "Ilość pytań: " + CStr(WierszNr - 1)
            Cells(WierszNr + 4, 1).Value = "Odp. dobre: " + CStr(OdpDobre)
            Cells(WierszNr + 5, 1).Value = "Odp. złe: " + CStr(OdpZle)
            'Oblicz procent.
                Procent = 0
                if(WierszNr > 1) then
                    Procent = (OdpDobre * 100 / (WierszNr - 1))
                End If
                Cells(WierszNr + 6, 1).Value = "Test zdany w: " + CStr(Procent) + "%"
            'Wystaw ocenę.
                JakaOcena = ""
                If (Procent > 95) Then
                    JakaOcena = "Bardzo dobry (5.0)"
                ElseIf (Procent > 80) Then
                        JakaOcena = "Dobry (4.0)"
                    ElseIf (Procent > 70) Then
                            JakaOcena = "Dostateczny (3.0)"
                        ElseIf (Procent > 50) Then
                                JakaOcena = "Dopuszczający (2.0)"
                            Else
                                JakaOcena = "Jedynka (1.0)"
                            End If
                Cells(WierszNr + 7, 1).Value = "Otrzymana ocena to: " + JakaOcena
        End If
End Sub