Sub WierszUsuwajPusteKomorki()
'WierszUsuwajPusteKomorki - Usuwaj puste komórki w kolumnie nr 1 (tj. kolumnie A).
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
        'Kopiuj dane z kolumny nr 2(B) do kolumny nr 1(A).
        W = 0
        For W = 1 To Z
            Cells(W, 1).Value = Trim(Cells(W, 2).Value)
            Cells(W, 2).Value = ""  'Usuń zawartość komórek.
        Next W
    Else
        MsgBox ("BŁĄD -?Nr wiersza musi być > 1!")
    End If
End Sub