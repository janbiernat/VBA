Sub ArkuszOdczytNazw()
'ArkuszOdczytNazw (c)by Jan T. Biernat
'
'Odczytanie nazw arkuszy z zewnętrznego zeszytu (tj. np. "zeszyt1.xlsx").
'
    Dim Wiersz As Integer
    Wiersz = 0
    With Workbooks("zeszyt1.xlsx")
        For Wiersz = 1 To .Worksheets.Count
            Sheets("Arkusz1").Cells(Wiersz, 1).Value = .Worksheets(Wiersz).Name 'Wpisz nazwy arkuszy w poszczególnych komórkach w kolumnie 1 (tj. w kolumnie A).
        Next Wiersz
    End With
End Sub