Sub WierszUsuwajPusteKomorki2()
'WierszUsuwajPusteKomorki2 - Usuwaj puste komórki w kolumnie nr 1 (tj. kolumnie A). Wersja 2.
'Copyright (c)by Jan T. Biernat
'
    Dim Z As Long
    Dim W As Long
    Dim L As Long
    Z = 0
    W = 0
    L = 0
    Z = Val(InputBox("Wpisz nr wiersza:", "Nr wiersza", ""))
    If (Z > 1) Then
        'Kopiuj dane z kolumny nr 1(A)
        'do kolumny nr 2(B) bez pustych komórek.
        For W = 1 To Z
            If (Trim(Cells(W, 1).Value) <> "") Then
                L = L + 1
                Cells(L, 2).Value = Trim(Cells(W, 1).Value)
            End If
        Next W
        'Usuń kolumnę nr 1 (tj. kolumnę A).
        Columns("A:A").Select
        Selection.Delete
        Range("B1").Select
    Else
        MsgBox ("BŁĄD -?Nr wiersza musi być > 1!")
    End If
End Sub