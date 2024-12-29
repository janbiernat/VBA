Function CzasGM_TerazJest() As String
'CzasGM_TerazJest - Funkcja zwraca bieżącą godzinę w formacie GG:MM.
'Copyright (c)by Jan T. Biernat
'
    Dim H As String
    Dim M As String
    H = ""
    H = CStr(Hour(Now()))
    M = ""
    M = CStr(Minute(Now()))
    If (Len(H) = 1) Then
        H = "0" + H
    End If
    If (Len(M) = 1) Then
        M = "0" + M
    End If
    CzasGM_TerazJest = H + ":" + M
End Function

Private Sub Workbook_Open()
    '
    'Umieść bieżący czas w komórce B1 (tj. w wierszu nr 1 i kolumnie nr 2), w arkuszu o nazwie "Arkusz1".
    With Sheets("Arkusz1")
        .Cells(1, 1).Value = "Czas:"
        .Cells(1, 2).Value = "'" + CzasGM_TerazJest()
    End With
End Sub