Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'Tabliczka mnożenia na krzyż z czyszczeniem komórek.
'Copyright (c)by Jan T. Biernat
'
    Dim Wiersz As Integer
    Dim Kolumna As Integer
    Dim A As Integer
    Dim B As Integer
    With Sheets("Tabliczka mnożenia na krzyż")
        .Cells(1, 1).Value = "Zakres:"
        .Cells(2, 4).Value = ""
        'Generuj tabliczkę mnożenia na krzyż
         If (.Cells(1, 2).Value <> "") Then
             If ((.Cells(1, 2).Value > 0) And (.Cells(1, 2).Value < 32)) Then
                 .Cells(1, 3).Value = "Wykonuję zadanie ..."
                 'Generuj tabliczkę mnożenia na krzyż.
                  For A = 1 To 31
                     If (A < .Cells(1, 2).Value + 1) Then
                        .Cells(3 + A, 1).Value = A
                        .Cells(3, 1 + A).Value = A
                     Else
                        .Cells(3 + A, 1).Value = ""
                        .Cells(3, 1 + A).Value = ""
                     End If
                     For B = 1 To 31
                        If ((B < .Cells(1, 2).Value + 1) And (A < .Cells(1, 2).Value + 1)) Then
                            .Cells(3 + A, 1 + B).Value = (A * B)
                        Else
                            .Cells(3 + A, 1 + B).Value = ""
                        End If
                     Next B
                  Next A
                 .Cells(1, 3).Value = "Jestem już gotowy!"
             End If
         Else
            .Cells(1, 3).Value = "BŁĄD -?Brak podanej liczby!"
            .Cells(2, 4).Value = "Proszę o podanie liczby z zakresu od 1 do 31."
         End If
    End With
End Sub