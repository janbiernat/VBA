Function fl12_345(T As String) As String
'fl12_345 - Funkcja formatuje liczbę (np. 12345 -> 12 345).
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
            S = Mid(T, I, 1) + S
            L = L + 1
            If (L > 2) Then
                S = " " + S
                L = 0
            End If
        Next I
        'Likwiduj znak spacji na początku ciągu znaków.
        If (Mid(S, 1, 1) = " ") Then
            fl12_345 = "'" + Mid(S, 2, Len(S))
        Else
            fl12_345 = "'" + S
        End If
    Else
        fl12_345 = "BŁĄD -?Brak danych!"
    End If
End Function

Sub FormatujLiczbe()
'Formatowanie liczb.
'
    Dim W As Integer
    W = 0
    Do
        W = W + 1
        Cells(W, 3).Value = fl12_345(Cells(W, 2).Value)
    Loop Until (Trim(Cells(W, 1).Value) = "")
End Sub