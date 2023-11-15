Function fl12_345(T As String) As String
'fl12_345 - Funkcja formatuje liczbę (tj. 12345 -> 12 345)
'Copyright (c)by Jan T. Biernat
'
    Dim I As Integer
    Dim L As Integer
    Dim S As String
    I = 0
    L = 0
    S = ""
    T = Trim(T)
    If (T <> "") Then
        For I = Len(T) To 1 Step -1
            L = L + 1
            S = Mid(T, I, 1) + S
            If (L > 2) Then
                L = 0
                S = " " + S
            End If
        Next I
    End If
    'Usuwa 1 znak pusty z podanego ciągu znaków.
    If (Mid(S, 1, 1) = " ") Then
        S = Mid(S, 2, Len(S))
    End If
    'Funkcja zwraca komunikat o braku danych, gdy zawartość zmiennej tekstowej "S" jest pusta.
    If (S <> "") Then
        fl12_345 = "'" + S
    Else
        fl12_345 = "BŁĄD -?Brak danych!"
    End If
End Function

Sub FormatujLiczbe()
'Formatuj liczby (np. 12345 na 12 345).
'
    Dim W As Integer
    W = 0
    Do
        W = W + 1
        Cells(W, 2).Value = fl12_345(Cells(W, 1).Value)
    Loop Until (Trim(Cells(W, 1).Value) = "")
    Cells(1, 2).Value = fl12_345(Cells(1, 1).Value)
End Sub