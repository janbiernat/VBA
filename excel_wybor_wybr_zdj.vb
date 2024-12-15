Sub SprWybor()
'SprWybor: Sprawdź poprawność wyboru zdjęć.
'Copyright (c)by Jan T. Biernat
'
    Dim WierszNr As Integer
    Dim WierszDodaj As Integer
    Dim SumaTrafien As Integer
    Dim TabDane(1) As String
    '
    WierszNr = 0
    WierszDodaj = 0
    SumaTrafien = 0
    '
    'Porównaj zawartość formularza z szablonem wzorcowym.
    With Sheets("Formularz")
        If (Trim(.Cells(4, 3).Value) <> "") Then
            Do
                WierszNr = WierszNr + 1
                TabDane(0) = ""
                TabDane(0) = LCase(Trim(.Cells(6 + WierszNr, 2).Value)) + CStr(Trim(.Cells(6 + WierszNr, 3).Value))
                TabDane(1) = ""
                TabDane(1) = LCase(Trim(Sheets("Opcje").Cells(28 + WierszNr, 2).Value)) + CStr(Trim(Sheets("Opcje").Cells(28 + WierszNr, 3).Value))
                If ((TabDane(0) <> "") And (TabDane(1) <> "")) Then
                    If (TabDane(0) = TabDane(1)) Then
                        SumaTrafien = SumaTrafien + 1   'Sumuj ilość trafień.
                    End If
                End If
            Loop Until (Trim(.Cells(6 + WierszNr, 2).Value) = "")
        Else
            MsgBox ("BŁĄD -?Brak oznaczenia klasy oraz imienia i nazwiska ucznia!")
        End If
    End With
    '
    'Raport: Umieść ucznia wraz z liczbą trafień.
    With Sheets("Raport")
        'Znajdź 1 pustą komórkę.
        Do
            WierszDodaj = WierszDodaj + 1
        Loop Until (Trim(.Cells(WierszDodaj + 3, 2).Value) = "")
        '
        'Wstaw kolejnego ucznia wraz z ilością trafień.
        .Cells(WierszDodaj + 3, 1).Value = WierszDodaj                                  'Kolejny nr ucznia.
        .Cells(WierszDodaj + 3, 2).Value = Trim(Sheets("Formularz").Cells(4, 3).Value)  'Klasa oraz imię i nazwisko ucznia.
        .Cells(WierszDodaj + 3, 3).Value = SumaTrafien                                  'Liczba trafień.
    End With
    '
    Range("C7:C28").Select  'Zaznacz zakres od C7 do C28.
    Selection.ClearContents 'Usuń zawartość zaznaczonych komórek (tj. od C7 do C28).
    Range("C7").Select      'Zaznacz komórkę C7.
End Sub