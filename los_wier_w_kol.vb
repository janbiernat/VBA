Function LosujWierWkol(NrKol As Integer, ArkuszNazwa As String) As String
'Losuj wiersz w wybranej kolumnie (c)by Jan T. Biernat
'
    'Deklaracja zmiennych.
    Dim W As Integer
    Dim LosW As Integer
    Dim Rezultat As String
    '
    'Ustawienia startowe
    W = 0
    LosW = 0
    Rezultat = ""
    '
    'Zabezpieczenie.
    If (NrKol < 1) Then
        NrKol = 1
    End If
    '
    'Oblicz ile komórek jest wypełnionych (licząc od 1 komórki) w wybranej kolumnie.
    With Sheets(ArkuszNazwa)
        Do
            W = W + 1
        Loop Until (Trim(.Cells(W, NrKol).Value) = "")
        '
        'Losowanie wiersza w wybranej kolumnie.
        Randomize
        LosW = 0
        LosW = Int(((W - 1) * Rnd) + 1)
        Rezultat = Trim(.Cells(LosW, NrKol).Value)  'Pobierz zawartość wylosowanej komórki.
    End With
    LosujWierWkol = Rezultat
End Function

Sub LosujWiersz()
'Losuj informację z komórki z wybranej kolumny.
'
    With Sheets("Arkusz1")
        .Cells(1, 4).Value = LosujWierWkol(0, "Arkusz1")
    End With
End Sub