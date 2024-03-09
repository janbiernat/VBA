Sub PunktyNaProcenty()
'PunktyNaProcenty - Funkcja przelicza punkty na procenty (%).
'Copyright (c)by Jan T. Biernat
'
    Const Info As String = "INFO -?W komórce A1 wpisz maksymalną ilość punktów."
    Dim I As Integer
    Dim Punkty As Integer
    I = 0
    Punkty = 0
    Cells(1, 1).Value = Trim(Cells(1, 1).Value)
    If (Cells(1, 1).Value <> "") Then
        Punkty = Val(Cells(1, 1).Value)
        If (Punkty > 0) Then
            Cells(2, 1).Value = "Punkty"
            Cells(2, 2).Value = "Procenty (%)"
            For I = 1 To Punkty
                Cells(I + 2, 1).Value = I
                Cells(I + 2, 2).NumberFormat = "0"
                Cells(I + 2, 2).Value = (Cells(I + 2, 1).Value * 100 / Punkty)
            Next I
        Else
            Cells(1, 1).Value = Info
        End If
    Else
        Cells(1, 1).Value = Info
    End If
    Cells(1, 2).Value = ""
End Sub