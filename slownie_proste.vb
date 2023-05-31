Function SlownieProste(Str As String, Waluta As String) As String
'Słownie proste (c)by Jan T. Biernat
'Funkcja zwraca skrócony opis podanej kwoty
'(np. 136zł = sto*trzy*sze*zł).
'
    Dim I As Integer
    Dim W As String
    I = 0
    W = ""
    Str = Trim(Str)
    If (Str <> "") Then
        For I = 1 To Len(Str)
            Select Case Mid(Str, I, 1)
                Case "0"            '0 - Zero.
                    W = W + "zer*"
                Case "1"            '1 - Jeden.
                    W = W + "jed*"
                Case "2"            '2 - Dwa.
                    W = W + "dwa*"
                Case "3"            '3 - Trzy.
                    W = W + "trz*"
                Case "4"            '4 - Cztery.
                    W = W + "czt*"
                Case "5"            '5 - Pięć.
                    W = W + "pię*"
                Case "6"            '6 - Sześć.
                    W = W + "sze*"
                Case "7"            '7 - Siedem.
                    W = W + "sie*"
                Case "8"            '8 - Osiem.
                    W = W + "osi*"
                Case "9"            '9 - Dziewięć.
                    W = W + "dzi*"
                Case Else
                    W = W + "?*"
            End Select
        Next I
        SlownieProste = W + Waluta
    Else
        SlownieProste = "BŁĄD -?Brak danych!"
    End If
End Function

Sub slownie_proste()
'Słownie proste – wywołanie.
'
    Dim I As Integer
    I = 0
    For I = 1 To 5
        Cells(I, 2).Value = SlownieProste(Cells(I, 1).Value, "zł")
    Next I
End Sub