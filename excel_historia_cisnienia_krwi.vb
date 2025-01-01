'
'Excel,VBA: Historia/Tabela ciśnienia krwi.
'Copyright (c)by Jan T. Biernat
'
Function DataRMD_TerazJest() As String
'DataRMD_TerazJest - Funkcja zwraca datę systemową (bieżącą) w postaci RRRR-MM-DD.
'
    Dim R As String
    Dim M As String
    Dim D As String
    R = ""
    R = CStr(Year(Date))
    M = ""
    M = CStr(Month(Date))
    D = ""
    D = CStr(Day(Date))
    If (Len(M) = 1) Then
        M = "0" + M
    End If
    If (Len(D) = 1) Then
        D = "0" + D
    End If
    DataRMD_TerazJest = R + "-" + M + "-" + D
End Function

Function CzasGM_TerazJest() As String
'CzasGM_TerazJest - Funkcja zwraca bieżącą godzinę w formacie GG:MM.
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

Sub FormDataCzasAktualizuj()
'FormDataCzasAktualizuj.
'Funkcja aktualizuje datę i czas bieżący.
'
    'Arkusz "Formularz": Ustawienia startowe.
    With Sheets("Formularz")
        .Cells(3, 4).Value = "'" + DataRMD_TerazJest()
        .Cells(4, 4).Value = "'" + CzasGM_TerazJest()
    End With
End Sub

Sub FormWyczysc()
'FormWyczysc:Czyści formularz.
'
    With Sheets("Formularz")
        .Cells(3, 4).Value = ""
        .Cells(4, 4).Value = ""
        .Cells(6, 4).Value = ""
        .Cells(7, 4).Value = ""
        .Cells(8, 4).Value = ""
        .Cells(9, 4).Value = ""
        Call FormDataCzasAktualizuj
        .Range("D6").Select
    End With
End Sub

Sub DaneWprowadz()
'DaneWprowadz: Wprowadzenie danych do tabeli ciśnienia krwi.
'
    Const WierszOd As Integer = 5   'Od którego wiersza wprowadzać dane
                                    '(Arkusz "Historia ciśnienia krwi").
    '
    Dim Wiersz As Integer
    Dim Tak As Boolean
    '
    'Arkusz "Formularz":
    'Sprawdź, czy wszystkie pola są uzupełnione.
    Tak = False
    With Sheets("Formularz")
        If ((Trim(.Cells(3, 4).Value) <> "") And (Trim(.Cells(4, 4).Value) <> "") And (Trim(.Cells(6, 4).Value) <> "") And (Trim(.Cells(7, 4).Value) <> "") And (Trim(.Cells(8, 4).Value) <> "")) Then
            Tak = True
        Else
            MsgBox ("BŁĄD -?Proszę uzupełnić wszystkie pola!")
        End If
    End With
    '
    'Arkusz "Historia ciśnienia krwi":
    'Umieszczenie danych z arkusza "Formularz".
    If (Tak = True) Then
        With Sheets("Historia ciśnienia krwi")
            'Znajdź 1 pustą komórkę.
            Wiersz = 0
            Do
                Wiersz = Wiersz + 1
            Loop Until (Trim(.Cells(Wiersz + WierszOd, 2).Value) = "")
            '
            'Kopiuj dane z formularza do historii ciśnienia krwi.
            .Cells(Wiersz + WierszOd, 1).Value = Wiersz                                            'LP.
            .Cells(Wiersz + WierszOd, 2).Value = "'" + Trim(Sheets("Formularz").Cells(3, 4).Value) 'Data badania.
            .Cells(Wiersz + WierszOd, 3).Value = "'" + Trim(Sheets("Formularz").Cells(4, 4).Value) 'Godzina badania.
            .Cells(Wiersz + WierszOd, 4).Value = Trim(Sheets("Formularz").Cells(6, 4).Value)       'Ciśnienie skurczowe.
            .Cells(Wiersz + WierszOd, 5).Value = Trim(Sheets("Formularz").Cells(7, 4).Value)       'Ciśnienie rozkurczowe.
            .Cells(Wiersz + WierszOd, 6).Value = Trim(Sheets("Formularz").Cells(8, 4).Value)       'Tętno.
            .Cells(Wiersz + WierszOd, 7).Value = Trim(Sheets("Formularz").Cells(9, 4).Value)       'Uwagi.
        End With
    End If
    Call FormWyczysc    'Wyczyść formularz.
End Sub